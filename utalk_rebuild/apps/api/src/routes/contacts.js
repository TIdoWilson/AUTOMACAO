import { randomUUID } from 'node:crypto'
import { Router } from 'express'
import { z } from 'zod'
import { query } from '../db/pool.js'
import { requireAuth } from '../middleware/auth.js'
import { writeAuditLog } from '../utils/audit.js'

export const contactsRouter = Router()

const createContactSchema = z.object({
  displayName: z.string().min(2),
  channelIdentifier: z.string().min(2),
  phoneMasked: z.string().optional(),
})

const updateContactSchema = z.object({
  displayName: z.string().min(2).optional(),
  phoneMasked: z.string().optional(),
  isActive: z.boolean().optional(),
})

contactsRouter.get('/', requireAuth, async (req, res, next) => {
  try {
    const search = (req.query.search ?? '').toString().trim().toLowerCase()
    const result = await query(
      `SELECT id, display_name AS "displayName", channel_identifier AS "channelIdentifier",
              phone_masked AS "phoneMasked", is_active AS "isActive", created_at AS "createdAt", updated_at AS "updatedAt"
       FROM contacts
       WHERE organization_id = $1
         AND ($2 = '' OR LOWER(display_name) LIKE '%' || $2 || '%' OR LOWER(channel_identifier) LIKE '%' || $2 || '%')
       ORDER BY updated_at DESC
       LIMIT 500`,
      [req.auth.organizationId, search],
    )
    return res.json({ contacts: result.rows })
  } catch (error) {
    return next(error)
  }
})

contactsRouter.post('/', requireAuth, async (req, res, next) => {
  try {
    const payload = createContactSchema.parse(req.body)
    const created = await query(
      `INSERT INTO contacts (id, organization_id, display_name, channel_identifier, phone_masked)
       VALUES ($1, $2, $3, $4, $5)
       RETURNING id, display_name AS "displayName", channel_identifier AS "channelIdentifier",
                 phone_masked AS "phoneMasked", is_active AS "isActive", created_at AS "createdAt", updated_at AS "updatedAt"`,
      [
        randomUUID(),
        req.auth.organizationId,
        payload.displayName.trim(),
        payload.channelIdentifier.trim(),
        payload.phoneMasked?.trim() ?? null,
      ],
    )

    const contact = created.rows[0]
    await writeAuditLog({
      organizationId: req.auth.organizationId,
      userId: req.auth.userId,
      action: 'contact_created',
      entityType: 'contact',
      entityId: contact.id,
    })
    return res.status(201).json({ contact })
  } catch (error) {
    return next(error)
  }
})

contactsRouter.patch('/:contactId', requireAuth, async (req, res, next) => {
  try {
    const payload = updateContactSchema.parse(req.body)
    const fields = []
    const values = []
    let idx = 1

    if (payload.displayName !== undefined) {
      fields.push(`display_name = $${idx++}`)
      values.push(payload.displayName.trim())
    }
    if (payload.phoneMasked !== undefined) {
      fields.push(`phone_masked = $${idx++}`)
      values.push(payload.phoneMasked.trim())
    }
    if (payload.isActive !== undefined) {
      fields.push(`is_active = $${idx++}`)
      values.push(payload.isActive)
    }

    if (fields.length === 0) {
      return res.status(400).json({ error: 'No valid fields to update.' })
    }

    fields.push(`updated_at = NOW()`)
    values.push(req.params.contactId, req.auth.organizationId)

    const updated = await query(
      `UPDATE contacts
       SET ${fields.join(', ')}
       WHERE id = $${idx++} AND organization_id = $${idx}
       RETURNING id, display_name AS "displayName", channel_identifier AS "channelIdentifier",
                 phone_masked AS "phoneMasked", is_active AS "isActive", created_at AS "createdAt", updated_at AS "updatedAt"`,
      values,
    )

    if (updated.rowCount === 0) {
      return res.status(404).json({ error: 'Contact not found.' })
    }

    await writeAuditLog({
      organizationId: req.auth.organizationId,
      userId: req.auth.userId,
      action: 'contact_updated',
      entityType: 'contact',
      entityId: req.params.contactId,
      metadata: payload,
    })

    return res.json({ contact: updated.rows[0] })
  } catch (error) {
    return next(error)
  }
})
