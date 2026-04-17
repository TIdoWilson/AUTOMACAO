import { env } from '../config/env.js'

export function requireInternalToken(req, res, next) {
  const provided = req.header('x-internal-token')
  if (!provided || provided !== env.INTERNAL_API_TOKEN) {
    return res.status(401).json({ error: 'Invalid internal token.' })
  }
  return next()
}
