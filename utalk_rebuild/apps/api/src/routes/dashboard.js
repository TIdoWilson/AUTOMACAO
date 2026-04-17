import { Router } from 'express'
import { query } from '../db/pool.js'
import { requireAuth } from '../middleware/auth.js'

export const dashboardRouter = Router()

dashboardRouter.get('/summary', requireAuth, async (req, res, next) => {
  try {
    const countsResult = await query(
      `SELECT
         COUNT(*)::int AS total_conversations,
         COUNT(*) FILTER (WHERE status = 'finalized')::int AS finalized_conversations,
         COUNT(*) FILTER (WHERE status = 'waiting_customer')::int AS waiting_customer_conversations,
         COUNT(*) FILTER (WHERE status = 'queued')::int AS queued_conversations,
         COUNT(*) FILTER (WHERE status = 'in_progress')::int AS in_progress_conversations
       FROM conversations
       WHERE organization_id = $1`,
      [req.auth.organizationId],
    )

    const contactsResult = await query(
      `SELECT COUNT(*)::int AS active_contacts
       FROM contacts
       WHERE organization_id = $1
         AND is_active = TRUE`,
      [req.auth.organizationId],
    )

    const avgResponseResult = await query(
      `WITH first_contact AS (
         SELECT conversation_id, MIN(created_at) AS first_contact_at
         FROM messages
         WHERE sender_type = 'contact'
         GROUP BY conversation_id
       ),
       first_operator AS (
         SELECT conversation_id, MIN(created_at) AS first_operator_at
         FROM messages
         WHERE sender_type = 'operator'
         GROUP BY conversation_id
       )
       SELECT AVG(EXTRACT(EPOCH FROM (fo.first_operator_at - fc.first_contact_at)))::numeric(10,2) AS avg_seconds
       FROM conversations c
       JOIN first_contact fc ON fc.conversation_id = c.id
       JOIN first_operator fo ON fo.conversation_id = c.id
       WHERE c.organization_id = $1`,
      [req.auth.organizationId],
    )

    const perDepartmentResult = await query(
      `SELECT d.id, d.name,
              COUNT(c.id)::int AS conversations
       FROM departments d
       LEFT JOIN conversations c
         ON c.department_id = d.id
        AND c.organization_id = d.organization_id
       WHERE d.organization_id = $1
       GROUP BY d.id
       ORDER BY d.name ASC`,
      [req.auth.organizationId],
    )

    const avgSeconds = Number(avgResponseResult.rows[0]?.avg_seconds ?? 0)
    const avgResponseMinutes = Number.isFinite(avgSeconds) ? Math.round((avgSeconds / 60) * 100) / 100 : 0

    return res.json({
      summary: {
        ...countsResult.rows[0],
        active_contacts: contactsResult.rows[0].active_contacts,
        average_first_response_minutes: avgResponseMinutes,
      },
      by_department: perDepartmentResult.rows,
    })
  } catch (error) {
    return next(error)
  }
})
