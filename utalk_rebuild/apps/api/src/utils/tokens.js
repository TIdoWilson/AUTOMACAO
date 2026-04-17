import crypto from 'node:crypto'
import jwt from 'jsonwebtoken'
import { env } from '../config/env.js'

export function createAccessToken(payload) {
  return jwt.sign(payload, env.JWT_ACCESS_SECRET, { expiresIn: env.ACCESS_TOKEN_TTL })
}

export function verifyAccessToken(token) {
  return jwt.verify(token, env.JWT_ACCESS_SECRET)
}

export function createRefreshToken() {
  return crypto.randomBytes(48).toString('base64url')
}

export function hashRefreshToken(refreshToken) {
  return crypto.createHash('sha256').update(refreshToken).digest('hex')
}

export function refreshExpirationDate() {
  const now = new Date()
  now.setDate(now.getDate() + env.REFRESH_TOKEN_TTL_DAYS)
  return now
}
