import { query } from '../db/pool.js'

export async function getUserDepartmentIds(userId) {
  const result = await query(
    `SELECT department_id
     FROM user_departments
     WHERE user_id = $1`,
    [userId],
  )
  return result.rows.map((row) => row.department_id)
}

export async function canUserAccessConversation({ userId, role, organizationId, conversationId }) {
  if (role === 'admin') {
    const adminAccess = await query(
      `SELECT 1
       FROM conversations
       WHERE id = $1 AND organization_id = $2`,
      [conversationId, organizationId],
    )
    return adminAccess.rowCount > 0
  }

  const departmentIds = await getUserDepartmentIds(userId)
  if (departmentIds.length === 0) {
    return false
  }

  const operatorAccess = await query(
    `SELECT 1
     FROM conversations c
     LEFT JOIN conversation_participants cp
       ON cp.conversation_id = c.id
      AND cp.user_id = $3
      AND cp.left_at IS NULL
     WHERE c.id = $1
       AND c.organization_id = $2
       AND (c.department_id = ANY($4::text[]) OR cp.id IS NOT NULL)`,
    [conversationId, organizationId, userId, departmentIds],
  )

  return operatorAccess.rowCount > 0
}
