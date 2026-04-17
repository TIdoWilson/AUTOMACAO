import { randomUUID } from 'node:crypto'
import { Router } from 'express'
import { z } from 'zod'
import { query } from '../db/pool.js'
import { requireAuth, requireRole } from '../middleware/auth.js'
import { writeAuditLog } from '../utils/audit.js'

export const departmentsRouter = Router()

const createDepartmentSchema = z.object({
  name: z.string().min(2),
})

departmentsRouter.get('/', requireAuth, async (req, res, next) => {
  try {
    const result = await query(
      `SELECT id, name, is_active, created_at
       FROM departments
       WHERE organization_id = $1
       ORDER BY name ASC`,
      [req.auth.organizationId],
    )
    return res.json({ departments: result.rows })
  } catch (error) {
    return next(error)
  }
})

departmentsRouter.post('/', requireAuth, requireRole('admin'), async (req, res, next) => {
  try {
    const payload = createDepartmentSchema.parse(req.body)
    const department = await query(
      `INSERT INTO departments (id, organization_id, name)
       VALUES ($1, $2, $3)
       RETURNING id, name, is_active, created_at`,
      [randomUUID(), req.auth.organizationId, payload.name.trim()],
    )

    const created = department.rows[0]
    await writeAuditLog({
      organizationId: req.auth.organizationId,
      userId: req.auth.userId,
      action: 'department_created',
      entityType: 'department',
      entityId: created.id,
      metadata: { name: created.name },
    })

    return res.status(201).json({ department: created })
  } catch (error) {
    return next(error)
  }
})
