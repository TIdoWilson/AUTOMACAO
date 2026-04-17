import { randomUUID } from 'node:crypto'
import { Router } from 'express'
import { z } from 'zod'
import { query, withTransaction } from '../db/pool.js'
import { requireAuth } from '../middleware/auth.js'
import { writeAuditLog } from '../utils/audit.js'
import { comparePassword, hashPassword } from '../utils/passwords.js'
import {
  createAccessToken,
  createRefreshToken,
  hashRefreshToken,
  refreshExpirationDate,
} from '../utils/tokens.js'
import { env } from '../config/env.js'

export const authRouter = Router()

const bootstrapSchema = z.object({
  organizationName: z.string().min(2),
  departmentName: z.string().min(2).default('Geral'),
  name: z.string().min(2),
  email: z.string().email(),
  password: z.string().min(8),
})

const loginSchema = z.object({
  email: z.string().email(),
  password: z.string().min(8),
})

function setRefreshCookie(res, refreshToken) {
  res.cookie('utalk_refresh_token', refreshToken, {
    httpOnly: true,
    secure: env.COOKIE_SECURE,
    sameSite: 'lax',
    maxAge: env.REFRESH_TOKEN_TTL_DAYS * 24 * 60 * 60 * 1000,
    path: '/',
  })
}

authRouter.post('/bootstrap-admin', async (req, res, next) => {
  try {
    const payload = bootstrapSchema.parse(req.body)
    const existingUsers = await query('SELECT COUNT(*)::int AS count FROM users')
    if (existingUsers.rows[0].count > 0) {
      return res.status(409).json({ error: 'Bootstrap can only be used when no users exist.' })
    }

    const normalizedEmail = payload.email.trim().toLowerCase()
    const passwordHash = await hashPassword(payload.password)

    const created = await withTransaction(async (client) => {
      const organizationId = randomUUID()
      const departmentId = randomUUID()
      const userId = randomUUID()

      await client.query(
        'INSERT INTO organizations (id, name) VALUES ($1, $2)',
        [organizationId, payload.organizationName.trim()],
      )
      await client.query(
        'INSERT INTO departments (id, organization_id, name) VALUES ($1, $2, $3)',
        [departmentId, organizationId, payload.departmentName.trim()],
      )
      await client.query(
        `INSERT INTO users (id, organization_id, department_default_id, name, email, password_hash, role)
         VALUES ($1, $2, $3, $4, $5, $6, 'admin')`,
        [userId, organizationId, departmentId, payload.name.trim(), normalizedEmail, passwordHash],
      )
      await client.query(
        'INSERT INTO user_departments (user_id, department_id) VALUES ($1, $2)',
        [userId, departmentId],
      )

      return { organizationId, userId, departmentId }
    })

    await writeAuditLog({
      organizationId: created.organizationId,
      userId: created.userId,
      action: 'bootstrap_admin',
      entityType: 'user',
      entityId: created.userId,
      metadata: { departmentId: created.departmentId },
    })

    return res.status(201).json({
      message: 'Admin bootstrap completed.',
      organizationId: created.organizationId,
      userId: created.userId,
      departmentId: created.departmentId,
    })
  } catch (error) {
    return next(error)
  }
})

authRouter.post('/login', async (req, res, next) => {
  try {
    const payload = loginSchema.parse(req.body)
    const normalizedEmail = payload.email.trim().toLowerCase()
    const userResult = await query(
      `SELECT id, organization_id, name, email, role, password_hash, is_active
       FROM users
       WHERE email = $1`,
      [normalizedEmail],
    )
    const user = userResult.rows[0]
    if (!user || !user.is_active) {
      return res.status(401).json({ error: 'Invalid credentials.' })
    }

    const validPassword = await comparePassword(payload.password, user.password_hash)
    if (!validPassword) {
      return res.status(401).json({ error: 'Invalid credentials.' })
    }

    const accessToken = createAccessToken({
      sub: user.id,
      role: user.role,
      organizationId: user.organization_id,
    })
    const refreshToken = createRefreshToken()
    const refreshTokenHash = hashRefreshToken(refreshToken)
    const expiresAt = refreshExpirationDate()

    await query(
      `INSERT INTO sessions (id, user_id, refresh_token_hash, expires_at)
       VALUES ($1, $2, $3, $4)`,
      [randomUUID(), user.id, refreshTokenHash, expiresAt],
    )

    await writeAuditLog({
      organizationId: user.organization_id,
      userId: user.id,
      action: 'login_success',
      entityType: 'session',
      entityId: null,
    })

    setRefreshCookie(res, refreshToken)
    return res.json({
      accessToken,
      user: {
        id: user.id,
        name: user.name,
        email: user.email,
        role: user.role,
        organizationId: user.organization_id,
      },
    })
  } catch (error) {
    return next(error)
  }
})

authRouter.post('/refresh', async (req, res, next) => {
  try {
    const refreshToken = req.cookies?.utalk_refresh_token
    if (!refreshToken) {
      return res.status(401).json({ error: 'Refresh token not found.' })
    }

    const refreshTokenHash = hashRefreshToken(refreshToken)
    const sessionResult = await query(
      `SELECT s.id, s.user_id, s.expires_at, s.revoked_at, u.organization_id, u.role, u.name, u.email, u.is_active
       FROM sessions s
       JOIN users u ON u.id = s.user_id
       WHERE s.refresh_token_hash = $1`,
      [refreshTokenHash],
    )

    const session = sessionResult.rows[0]
    if (!session || session.revoked_at || new Date(session.expires_at) < new Date() || !session.is_active) {
      return res.status(401).json({ error: 'Refresh token is invalid or expired.' })
    }

    const newRefreshToken = createRefreshToken()
    const newRefreshTokenHash = hashRefreshToken(newRefreshToken)
    const newExpiration = refreshExpirationDate()

    await query(
      `UPDATE sessions
       SET refresh_token_hash = $1, expires_at = $2, updated_at = NOW()
       WHERE id = $3`,
      [newRefreshTokenHash, newExpiration, session.id],
    )

    const accessToken = createAccessToken({
      sub: session.user_id,
      role: session.role,
      organizationId: session.organization_id,
    })

    setRefreshCookie(res, newRefreshToken)
    return res.json({
      accessToken,
      user: {
        id: session.user_id,
        name: session.name,
        email: session.email,
        role: session.role,
        organizationId: session.organization_id,
      },
    })
  } catch (error) {
    return next(error)
  }
})

authRouter.post('/logout', requireAuth, async (req, res, next) => {
  try {
    const refreshToken = req.cookies?.utalk_refresh_token
    if (refreshToken) {
      const refreshTokenHash = hashRefreshToken(refreshToken)
      await query(
        `UPDATE sessions
         SET revoked_at = NOW(), updated_at = NOW()
         WHERE refresh_token_hash = $1`,
        [refreshTokenHash],
      )
    }
    res.clearCookie('utalk_refresh_token', { path: '/' })
    await writeAuditLog({
      organizationId: req.auth.organizationId,
      userId: req.auth.userId,
      action: 'logout',
      entityType: 'session',
      entityId: null,
    })
    return res.json({ message: 'Logged out.' })
  } catch (error) {
    return next(error)
  }
})

authRouter.get('/me', requireAuth, async (req, res) => {
  return res.json({
    user: {
      id: req.auth.userId,
      name: req.auth.name,
      email: req.auth.email,
      role: req.auth.role,
      organizationId: req.auth.organizationId,
    },
  })
})
