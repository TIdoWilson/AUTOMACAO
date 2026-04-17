import { randomUUID } from 'node:crypto'
import { Router } from 'express'
import { z } from 'zod'
import { query, withTransaction } from '../db/pool.js'
import { requireAuth, requireRole } from '../middleware/auth.js'
import { writeAuditLog } from '../utils/audit.js'
import { hashPassword } from '../utils/passwords.js'

export const usersRouter = Router()

const createUserSchema = z.object({
  name: z.string().min(2),
  email: z.string().email(),
  password: z.string().min(8),
  role: z.enum(['admin', 'operator']),
  departmentIds: z.array(z.string().uuid()).default([]),
})

usersRouter.get('/', requireAuth, requireRole('admin'), async (req, res, next) => {
  try {
    const result = await query(
      `SELECT u.id, u.name, u.email, u.role, u.is_active, u.created_at,
              COALESCE(
                JSON_AGG(
                  JSON_BUILD_OBJECT('id', d.id, 'name', d.name)
                ) FILTER (WHERE d.id IS NOT NULL),
                '[]'::json
              ) AS departments
       FROM users u
       LEFT JOIN user_departments ud ON ud.user_id = u.id
       LEFT JOIN departments d ON d.id = ud.department_id
       WHERE u.organization_id = $1
       GROUP BY u.id
       ORDER BY u.created_at DESC`,
      [req.auth.organizationId],
    )
    return res.json({ users: result.rows })
  } catch (error) {
    return next(error)
  }
})

usersRouter.post('/', requireAuth, requireRole('admin'), async (req, res, next) => {
  try {
    const payload = createUserSchema.parse(req.body)
    const normalizedEmail = payload.email.trim().toLowerCase()
    const passwordHash = await hashPassword(payload.password)

    const created = await withTransaction(async (client) => {
      const userId = randomUUID()
      const userResult = await client.query(
        `INSERT INTO users (id, organization_id, name, email, password_hash, role)
         VALUES ($1, $2, $3, $4, $5, $6)
         RETURNING id, name, email, role, is_active, created_at`,
        [userId, req.auth.organizationId, payload.name.trim(), normalizedEmail, passwordHash, payload.role],
      )

      if (payload.departmentIds.length > 0) {
        const allowedDepartments = await client.query(
          `SELECT id
           FROM departments
           WHERE organization_id = $1
             AND id = ANY($2::text[])`,
          [req.auth.organizationId, payload.departmentIds],
        )
        const validIds = new Set(allowedDepartments.rows.map((row) => row.id))
        const safeDepartmentIds = payload.departmentIds.filter((departmentId) => validIds.has(departmentId))
        for (const departmentId of safeDepartmentIds) {
          await client.query(
            `INSERT INTO user_departments (user_id, department_id)
             VALUES ($1, $2)
             ON CONFLICT DO NOTHING`,
            [userId, departmentId],
          )
        }
      }

      return userResult.rows[0]
    })

    await writeAuditLog({
      organizationId: req.auth.organizationId,
      userId: req.auth.userId,
      action: 'user_created',
      entityType: 'user',
      entityId: created.id,
      metadata: { role: created.role, email: created.email },
    })

    return res.status(201).json({ user: created })
  } catch (error) {
    return next(error)
  }
})
