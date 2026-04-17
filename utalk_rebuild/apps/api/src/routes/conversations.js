import { randomUUID } from 'node:crypto'
import { Router } from 'express'
import { z } from 'zod'
import { query, withTransaction } from '../db/pool.js'
import { requireAuth } from '../middleware/auth.js'
import { requireInternalToken } from '../middleware/internalToken.js'
import { canUserAccessConversation, getUserDepartmentIds } from '../services/accessService.js'
import { writeAuditLog } from '../utils/audit.js'

export const conversationsRouter = Router()

const incomingMessageSchema = z.object({
  organizationId: z.string().uuid(),
  channelIdentifier: z.string().min(2),
  contactName: z.string().min(1),
  phoneMasked: z.string().optional(),
  text: z.string().min(1),
})

const operatorMessageSchema = z.object({
  body: z.string().min(1),
  metadata: z.record(z.string(), z.any()).optional(),
})

const transferSchema = z.object({
  departmentId: z.string().uuid(),
})

function buildDepartmentMenuMessage(options) {
  const lines = options.map((item, index) => `${index + 1} - ${item.optionLabel}`).join('\n')
  return `Selecione o departamento desejado respondendo com o numero da opcao:\n${lines}`
}

function resolveDepartmentByText(text, options) {
  const match = String(text).trim().match(/^(\d+)$/)
  if (!match) return null
  const selectedIndex = Number(match[1]) - 1
  if (selectedIndex < 0 || selectedIndex >= options.length) return null
  return options[selectedIndex]
}

async function getOrCreateTriageFlow(client, organizationId) {
  const flowName = 'Default department triage'
  const flowResult = await client.query(
    `SELECT id
     FROM chatbot_flows
     WHERE organization_id = $1
       AND name = $2`,
    [organizationId, flowName],
  )

  let flowId = flowResult.rows[0]?.id
  if (!flowId) {
    flowId = randomUUID()
    await client.query(
      `INSERT INTO chatbot_flows (id, organization_id, name, is_active)
       VALUES ($1, $2, $3, TRUE)`,
      [flowId, organizationId, flowName],
    )
  }

  const departmentsResult = await client.query(
    `SELECT id, name
     FROM departments
     WHERE organization_id = $1
       AND is_active = TRUE
     ORDER BY name ASC`,
    [organizationId],
  )
  const departments = departmentsResult.rows

  for (let i = 0; i < departments.length; i += 1) {
    const department = departments[i]
    const exists = await client.query(
      `SELECT 1
       FROM chatbot_flow_departments
       WHERE flow_id = $1
         AND department_id = $2`,
      [flowId, department.id],
    )
    if (exists.rowCount === 0) {
      await client.query(
        `INSERT INTO chatbot_flow_departments (id, flow_id, department_id, option_order, option_label)
         VALUES ($1, $2, $3, $4, $5)`,
        [randomUUID(), flowId, department.id, i + 1, department.name],
      )
    }
  }

  const optionsResult = await client.query(
    `SELECT cfd.department_id AS "departmentId", cfd.option_label AS "optionLabel", d.name AS "departmentName", cfd.option_order AS "optionOrder"
     FROM chatbot_flow_departments cfd
     JOIN departments d ON d.id = cfd.department_id
     WHERE cfd.flow_id = $1
       AND d.is_active = TRUE
     ORDER BY cfd.option_order ASC, cfd.option_label ASC`,
    [flowId],
  )

  return {
    flowId,
    options: optionsResult.rows,
  }
}

async function insertMessage(client, { conversationId, senderType, senderUserId = null, body, metadata = {} }) {
  const messageId = randomUUID()
  await client.query(
    `INSERT INTO messages (id, conversation_id, sender_type, sender_user_id, body, metadata)
     VALUES ($1, $2, $3, $4, $5, $6::jsonb)`,
    [messageId, conversationId, senderType, senderUserId, body, JSON.stringify(metadata)],
  )
  return messageId
}

conversationsRouter.post('/incoming-message', requireInternalToken, async (req, res, next) => {
  try {
    const payload = incomingMessageSchema.parse(req.body)

    const result = await withTransaction(async (client) => {
      let contactId
      const contactLookup = await client.query(
        `SELECT id
         FROM contacts
         WHERE organization_id = $1
           AND channel_identifier = $2`,
        [payload.organizationId, payload.channelIdentifier.trim()],
      )

      if (contactLookup.rowCount === 0) {
        contactId = randomUUID()
        await client.query(
          `INSERT INTO contacts (id, organization_id, display_name, channel_identifier, phone_masked)
           VALUES ($1, $2, $3, $4, $5)`,
          [
            contactId,
            payload.organizationId,
            payload.contactName.trim(),
            payload.channelIdentifier.trim(),
            payload.phoneMasked?.trim() ?? null,
          ],
        )
      } else {
        contactId = contactLookup.rows[0].id
      }

      const latestConversation = await client.query(
        `SELECT id, status, department_id AS "departmentId", last_department_id AS "lastDepartmentId"
         FROM conversations
         WHERE organization_id = $1
           AND contact_id = $2
         ORDER BY opened_at DESC
         LIMIT 1`,
        [payload.organizationId, contactId],
      )

      let conversation = latestConversation.rows[0]

      if (!conversation || conversation.status === 'finalized') {
        const conversationId = randomUUID()
        await client.query(
          `INSERT INTO conversations (id, organization_id, contact_id, status, department_id, last_department_id)
           VALUES ($1, $2, $3, 'bot_triage', NULL, $4)`,
          [conversationId, payload.organizationId, contactId, conversation?.lastDepartmentId ?? null],
        )
        conversation = {
          id: conversationId,
          status: 'bot_triage',
          departmentId: null,
          lastDepartmentId: conversation?.lastDepartmentId ?? null,
        }
      }

      await insertMessage(client, {
        conversationId: conversation.id,
        senderType: 'contact',
        body: payload.text.trim(),
      })

      const activeParticipants = await client.query(
        `SELECT COUNT(*)::int AS count
         FROM conversation_participants
         WHERE conversation_id = $1
           AND left_at IS NULL`,
        [conversation.id],
      )
      const activeParticipantCount = activeParticipants.rows[0].count

      let triageOptions = []
      let statusAfter = conversation.status
      let departmentAfter = conversation.departmentId

      if (conversation.status === 'bot_triage') {
        const triage = await getOrCreateTriageFlow(client, payload.organizationId)
        triageOptions = triage.options

        if (triageOptions.length === 0) {
          await insertMessage(client, {
            conversationId: conversation.id,
            senderType: 'bot',
            body: 'Nenhum departamento ativo foi configurado. Avise um administrador.',
          })
        } else {
          const selected = resolveDepartmentByText(payload.text, triageOptions)
          if (!selected) {
            await insertMessage(client, {
              conversationId: conversation.id,
              senderType: 'bot',
              body: buildDepartmentMenuMessage(triageOptions),
            })

            await client.query(
              `INSERT INTO chatbot_executions (id, organization_id, conversation_id, flow_id, stage, is_completed)
               VALUES ($1, $2, $3, $4, 'select_department', FALSE)`,
              [randomUUID(), payload.organizationId, conversation.id, triage.flowId],
            )
          } else {
            await client.query(
              `UPDATE conversations
               SET status = 'queued',
                   department_id = $1,
                   last_department_id = $1,
                   updated_at = NOW()
               WHERE id = $2`,
              [selected.departmentId, conversation.id],
            )
            statusAfter = 'queued'
            departmentAfter = selected.departmentId

            await client.query(
              `INSERT INTO chatbot_executions (id, organization_id, conversation_id, flow_id, stage, is_completed, updated_at)
               VALUES ($1, $2, $3, $4, 'department_selected', TRUE, NOW())`,
              [randomUUID(), payload.organizationId, conversation.id, triage.flowId],
            )

            await insertMessage(client, {
              conversationId: conversation.id,
              senderType: 'bot',
              body: `Perfeito. Encaminhei voce para o departamento ${selected.departmentName}. Em instantes um atendente assumira.`,
              metadata: { departmentId: selected.departmentId },
            })
          }
        }
      } else if (activeParticipantCount === 0 && conversation.lastDepartmentId) {
        await client.query(
          `UPDATE conversations
           SET status = 'queued',
               department_id = last_department_id,
               updated_at = NOW()
           WHERE id = $1`,
          [conversation.id],
        )
        statusAfter = 'queued'
        departmentAfter = conversation.lastDepartmentId

        await insertMessage(client, {
          conversationId: conversation.id,
          senderType: 'system',
          body: 'Cliente retornou e a conversa voltou para a fila do ultimo departamento.',
        })
      } else if (conversation.status === 'waiting_customer' && activeParticipantCount > 0) {
        await client.query(
          `UPDATE conversations
           SET status = 'in_progress',
               updated_at = NOW()
           WHERE id = $1`,
          [conversation.id],
        )
        statusAfter = 'in_progress'
      }

      return {
        conversationId: conversation.id,
        contactId,
        status: statusAfter,
        departmentId: departmentAfter,
        triageOptions,
      }
    })

    await writeAuditLog({
      organizationId: payload.organizationId,
      action: 'incoming_message_processed',
      entityType: 'conversation',
      entityId: result.conversationId,
      metadata: { status: result.status, departmentId: result.departmentId },
    })

    return res.status(201).json(result)
  } catch (error) {
    return next(error)
  }
})

conversationsRouter.get('/', requireAuth, async (req, res, next) => {
  try {
    const statusFilter = (req.query.status ?? '').toString().trim()
    const onlyMine = (req.query.onlyMine ?? '').toString().toLowerCase() === 'true'
    const departmentFilter = (req.query.departmentId ?? '').toString().trim()
    const search = (req.query.search ?? '').toString().trim().toLowerCase()

    let allowedDepartmentIds = []
    if (req.auth.role === 'operator') {
      allowedDepartmentIds = await getUserDepartmentIds(req.auth.userId)
      if (allowedDepartmentIds.length === 0) {
        return res.json({ conversations: [] })
      }
    }

    const baseResult = await query(
      `SELECT c.id, c.status, c.department_id AS "departmentId", c.last_department_id AS "lastDepartmentId", c.updated_at AS "updatedAt",
              c.opened_at AS "openedAt", c.finalized_at AS "finalizedAt",
              ct.display_name AS "contactName", ct.channel_identifier AS "channelIdentifier",
              d.name AS "departmentName",
              EXISTS (
                SELECT 1
                FROM conversation_participants cp
                WHERE cp.conversation_id = c.id
                  AND cp.user_id = $2
                  AND cp.left_at IS NULL
              ) AS "isMine",
              (
                SELECT COUNT(*)::int
                FROM conversation_participants cp2
                WHERE cp2.conversation_id = c.id
                  AND cp2.left_at IS NULL
              ) AS "activeParticipants",
              (
                SELECT m.body
                FROM messages m
                WHERE m.conversation_id = c.id
                ORDER BY m.created_at DESC
                LIMIT 1
              ) AS "lastMessage"
       FROM conversations c
       JOIN contacts ct ON ct.id = c.contact_id
       LEFT JOIN departments d ON d.id = c.department_id
       WHERE c.organization_id = $1
       ORDER BY c.updated_at DESC
       LIMIT 500`,
      [req.auth.organizationId, req.auth.userId],
    )

    const filtered = baseResult.rows.filter((row) => {
      if (statusFilter && row.status !== statusFilter) return false
      if (departmentFilter && row.departmentId !== departmentFilter) return false
      if (onlyMine && !row.isMine) return false
      if (search) {
        const text = `${row.contactName} ${row.channelIdentifier} ${row.lastMessage ?? ''}`.toLowerCase()
        if (!text.includes(search)) return false
      }

      if (req.auth.role === 'operator') {
        const canSeeByDepartment = row.departmentId && allowedDepartmentIds.includes(row.departmentId)
        if (!canSeeByDepartment && !row.isMine) return false
      }
      return true
    })

    return res.json({ conversations: filtered })
  } catch (error) {
    return next(error)
  }
})

conversationsRouter.get('/:conversationId/messages', requireAuth, async (req, res, next) => {
  try {
    const allowed = await canUserAccessConversation({
      userId: req.auth.userId,
      role: req.auth.role,
      organizationId: req.auth.organizationId,
      conversationId: req.params.conversationId,
    })
    if (!allowed) {
      return res.status(403).json({ error: 'You cannot access this conversation.' })
    }

    const result = await query(
      `SELECT id, sender_type AS "senderType", sender_user_id AS "senderUserId", body, metadata, created_at AS "createdAt"
       FROM messages
       WHERE conversation_id = $1
       ORDER BY created_at ASC`,
      [req.params.conversationId],
    )
    return res.json({ messages: result.rows })
  } catch (error) {
    return next(error)
  }
})

conversationsRouter.post('/:conversationId/assume', requireAuth, async (req, res, next) => {
  try {
    const allowed = await canUserAccessConversation({
      userId: req.auth.userId,
      role: req.auth.role,
      organizationId: req.auth.organizationId,
      conversationId: req.params.conversationId,
    })
    if (!allowed) {
      return res.status(403).json({ error: 'You cannot assume this conversation.' })
    }

    const updated = await withTransaction(async (client) => {
      const conversationResult = await client.query(
        `SELECT id, status
         FROM conversations
         WHERE id = $1
           AND organization_id = $2`,
        [req.params.conversationId, req.auth.organizationId],
      )
      const conversation = conversationResult.rows[0]
      if (!conversation) return null
      if (conversation.status === 'finalized') {
        throw new Error('Cannot assume a finalized conversation.')
      }

      const existingParticipant = await client.query(
        `SELECT id
         FROM conversation_participants
         WHERE conversation_id = $1
           AND user_id = $2
           AND left_at IS NULL`,
        [req.params.conversationId, req.auth.userId],
      )
      if (existingParticipant.rowCount === 0) {
        await client.query(
          `INSERT INTO conversation_participants (id, conversation_id, user_id)
           VALUES ($1, $2, $3)`,
          [randomUUID(), req.params.conversationId, req.auth.userId],
        )
      }

      await client.query(
        `UPDATE conversations
         SET status = 'in_progress',
             updated_at = NOW()
         WHERE id = $1`,
        [req.params.conversationId],
      )

      await insertMessage(client, {
        conversationId: req.params.conversationId,
        senderType: 'system',
        body: `${req.auth.name} assumiu o atendimento.`,
      })

      return { id: req.params.conversationId, status: 'in_progress' }
    })

    if (!updated) {
      return res.status(404).json({ error: 'Conversation not found.' })
    }

    await writeAuditLog({
      organizationId: req.auth.organizationId,
      userId: req.auth.userId,
      action: 'conversation_assumed',
      entityType: 'conversation',
      entityId: req.params.conversationId,
    })

    return res.json({ conversation: updated })
  } catch (error) {
    return next(error)
  }
})

conversationsRouter.post('/:conversationId/leave', requireAuth, async (req, res, next) => {
  try {
    const allowed = await canUserAccessConversation({
      userId: req.auth.userId,
      role: req.auth.role,
      organizationId: req.auth.organizationId,
      conversationId: req.params.conversationId,
    })
    if (!allowed) {
      return res.status(403).json({ error: 'You cannot leave this conversation.' })
    }

    const result = await withTransaction(async (client) => {
      const leaveResult = await client.query(
        `UPDATE conversation_participants
         SET left_at = NOW()
         WHERE id = (
           SELECT id
           FROM conversation_participants
           WHERE conversation_id = $1
             AND user_id = $2
             AND left_at IS NULL
           ORDER BY joined_at DESC
           LIMIT 1
         )
         RETURNING id`,
        [req.params.conversationId, req.auth.userId],
      )
      if (leaveResult.rowCount === 0) {
        throw new Error('You are not an active participant of this conversation.')
      }

      const activeCountResult = await client.query(
        `SELECT COUNT(*)::int AS count
         FROM conversation_participants
         WHERE conversation_id = $1
           AND left_at IS NULL`,
        [req.params.conversationId],
      )
      const activeCount = activeCountResult.rows[0].count

      if (activeCount === 0) {
        await client.query(
          `UPDATE conversations
           SET status = CASE WHEN status = 'finalized' THEN 'finalized' ELSE 'waiting_customer' END,
               updated_at = NOW()
           WHERE id = $1`,
          [req.params.conversationId],
        )
      }

      await insertMessage(client, {
        conversationId: req.params.conversationId,
        senderType: 'system',
        body: `${req.auth.name} saiu da conversa.`,
      })

      return { activeParticipants: activeCount, status: activeCount === 0 ? 'waiting_customer' : 'in_progress' }
    })

    await writeAuditLog({
      organizationId: req.auth.organizationId,
      userId: req.auth.userId,
      action: 'conversation_left',
      entityType: 'conversation',
      entityId: req.params.conversationId,
      metadata: result,
    })

    return res.json(result)
  } catch (error) {
    return next(error)
  }
})

conversationsRouter.post('/:conversationId/finalize', requireAuth, async (req, res, next) => {
  try {
    const allowed = await canUserAccessConversation({
      userId: req.auth.userId,
      role: req.auth.role,
      organizationId: req.auth.organizationId,
      conversationId: req.params.conversationId,
    })
    if (!allowed) {
      return res.status(403).json({ error: 'You cannot finalize this conversation.' })
    }

    const updated = await withTransaction(async (client) => {
      const finalizeResult = await client.query(
        `UPDATE conversations
         SET status = 'finalized',
             finalized_at = NOW(),
             updated_at = NOW()
         WHERE id = $1
           AND organization_id = $2
         RETURNING id, status, finalized_at AS "finalizedAt"`,
        [req.params.conversationId, req.auth.organizationId],
      )
      if (finalizeResult.rowCount === 0) return null

      await client.query(
        `UPDATE conversation_participants
         SET left_at = NOW()
         WHERE conversation_id = $1
           AND left_at IS NULL`,
        [req.params.conversationId],
      )

      await insertMessage(client, {
        conversationId: req.params.conversationId,
        senderType: 'system',
        body: `${req.auth.name} finalizou a conversa.`,
      })

      return finalizeResult.rows[0]
    })

    if (!updated) {
      return res.status(404).json({ error: 'Conversation not found.' })
    }

    await writeAuditLog({
      organizationId: req.auth.organizationId,
      userId: req.auth.userId,
      action: 'conversation_finalized',
      entityType: 'conversation',
      entityId: req.params.conversationId,
    })

    return res.json({ conversation: updated })
  } catch (error) {
    return next(error)
  }
})

conversationsRouter.post('/:conversationId/waiting-customer', requireAuth, async (req, res, next) => {
  try {
    const allowed = await canUserAccessConversation({
      userId: req.auth.userId,
      role: req.auth.role,
      organizationId: req.auth.organizationId,
      conversationId: req.params.conversationId,
    })
    if (!allowed) {
      return res.status(403).json({ error: 'You cannot update this conversation.' })
    }

    const updated = await query(
      `UPDATE conversations
       SET status = 'waiting_customer',
           updated_at = NOW()
       WHERE id = $1
         AND organization_id = $2
       RETURNING id, status`,
      [req.params.conversationId, req.auth.organizationId],
    )
    if (updated.rowCount === 0) {
      return res.status(404).json({ error: 'Conversation not found.' })
    }

    await writeAuditLog({
      organizationId: req.auth.organizationId,
      userId: req.auth.userId,
      action: 'conversation_waiting_customer',
      entityType: 'conversation',
      entityId: req.params.conversationId,
    })

    return res.json({ conversation: updated.rows[0] })
  } catch (error) {
    return next(error)
  }
})

conversationsRouter.post('/:conversationId/transfer', requireAuth, async (req, res, next) => {
  try {
    const payload = transferSchema.parse(req.body)
    const allowed = await canUserAccessConversation({
      userId: req.auth.userId,
      role: req.auth.role,
      organizationId: req.auth.organizationId,
      conversationId: req.params.conversationId,
    })
    if (!allowed) {
      return res.status(403).json({ error: 'You cannot transfer this conversation.' })
    }

    const department = await query(
      `SELECT id, name
       FROM departments
       WHERE id = $1
         AND organization_id = $2
         AND is_active = TRUE`,
      [payload.departmentId, req.auth.organizationId],
    )
    if (department.rowCount === 0) {
      return res.status(404).json({ error: 'Department not found.' })
    }

    const departmentName = department.rows[0].name

    const updated = await withTransaction(async (client) => {
      const transferResult = await client.query(
        `UPDATE conversations
         SET department_id = $1,
             last_department_id = $1,
             status = 'queued',
             updated_at = NOW()
         WHERE id = $2
           AND organization_id = $3
         RETURNING id, status, department_id AS "departmentId"`,
        [payload.departmentId, req.params.conversationId, req.auth.organizationId],
      )
      if (transferResult.rowCount === 0) return null

      await insertMessage(client, {
        conversationId: req.params.conversationId,
        senderType: 'system',
        body: `${req.auth.name} transferiu para o departamento ${departmentName}.`,
        metadata: { departmentId: payload.departmentId },
      })

      return transferResult.rows[0]
    })

    if (!updated) {
      return res.status(404).json({ error: 'Conversation not found.' })
    }

    await writeAuditLog({
      organizationId: req.auth.organizationId,
      userId: req.auth.userId,
      action: 'conversation_transferred',
      entityType: 'conversation',
      entityId: req.params.conversationId,
      metadata: { departmentId: payload.departmentId },
    })

    return res.json({ conversation: updated })
  } catch (error) {
    return next(error)
  }
})

conversationsRouter.post('/:conversationId/messages', requireAuth, async (req, res, next) => {
  try {
    const payload = operatorMessageSchema.parse(req.body)
    const allowed = await canUserAccessConversation({
      userId: req.auth.userId,
      role: req.auth.role,
      organizationId: req.auth.organizationId,
      conversationId: req.params.conversationId,
    })
    if (!allowed) {
      return res.status(403).json({ error: 'You cannot send messages in this conversation.' })
    }

    const created = await withTransaction(async (client) => {
      const conversation = await client.query(
        `SELECT id, status
         FROM conversations
         WHERE id = $1
           AND organization_id = $2`,
        [req.params.conversationId, req.auth.organizationId],
      )
      if (conversation.rowCount === 0) return null
      if (conversation.rows[0].status === 'finalized') {
        throw new Error('Cannot send message in a finalized conversation.')
      }

      const participant = await client.query(
        `SELECT id
         FROM conversation_participants
         WHERE conversation_id = $1
           AND user_id = $2
           AND left_at IS NULL`,
        [req.params.conversationId, req.auth.userId],
      )
      if (participant.rowCount === 0) {
        await client.query(
          `INSERT INTO conversation_participants (id, conversation_id, user_id)
           VALUES ($1, $2, $3)`,
          [randomUUID(), req.params.conversationId, req.auth.userId],
        )
      }

      const messageId = await insertMessage(client, {
        conversationId: req.params.conversationId,
        senderType: 'operator',
        senderUserId: req.auth.userId,
        body: payload.body.trim(),
        metadata: payload.metadata ?? {},
      })

      await client.query(
        `UPDATE conversations
         SET status = 'in_progress',
             updated_at = NOW()
         WHERE id = $1`,
        [req.params.conversationId],
      )

      return { id: messageId, body: payload.body.trim(), senderType: 'operator' }
    })

    if (!created) {
      return res.status(404).json({ error: 'Conversation not found.' })
    }

    await writeAuditLog({
      organizationId: req.auth.organizationId,
      userId: req.auth.userId,
      action: 'operator_message_sent',
      entityType: 'conversation',
      entityId: req.params.conversationId,
    })

    return res.status(201).json({ message: created })
  } catch (error) {
    return next(error)
  }
})
