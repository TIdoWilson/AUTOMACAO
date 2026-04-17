import { query } from '../db/pool.js'
import { verifyAccessToken } from '../utils/tokens.js'

export async function requireAuth(req, res, next) {
  try {
    const authHeader = req.headers.authorization ?? ''
    const token = authHeader.startsWith('Bearer ') ? authHeader.slice('Bearer '.length) : null

    if (!token) {
      return res.status(401).json({ error: 'Authentication token not provided.' })
    }

    const decoded = verifyAccessToken(token)
    const userResult = await query(
      `SELECT id, organization_id, name, email, role, is_active
       FROM users
       WHERE id = $1`,
      [decoded.sub],
    )
    const user = userResult.rows[0]

    if (!user || !user.is_active) {
      return res.status(401).json({ error: 'User is not active or does not exist.' })
    }

    req.auth = {
      userId: user.id,
      organizationId: user.organization_id,
      role: user.role,
      name: user.name,
      email: user.email,
    }

    return next()
  } catch (error) {
    return res.status(401).json({ error: 'Invalid or expired token.' })
  }
}

export function requireRole(...allowedRoles) {
  return (req, res, next) => {
    if (!req.auth) {
      return res.status(401).json({ error: 'Authentication is required.' })
    }
    if (!allowedRoles.includes(req.auth.role)) {
      return res.status(403).json({ error: 'Insufficient permission.' })
    }
    return next()
  }
}
