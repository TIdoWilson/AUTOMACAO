import { randomUUID } from 'node:crypto'
import { query } from '../db/pool.js'

export async function writeAuditLog({
  organizationId = null,
  userId = null,
  action,
  entityType,
  entityId = null,
  metadata = {},
}) {
  await query(
    `INSERT INTO audit_logs (id, organization_id, user_id, action, entity_type, entity_id, metadata)
     VALUES ($1, $2, $3, $4, $5, $6, $7::jsonb)`,
    [randomUUID(), organizationId, userId, action, entityType, entityId, JSON.stringify(metadata)],
  )
}
