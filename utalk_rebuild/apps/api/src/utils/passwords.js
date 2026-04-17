import bcrypt from 'bcryptjs'

const SALT_ROUNDS = 12

export function hashPassword(rawPassword) {
  return bcrypt.hash(rawPassword, SALT_ROUNDS)
}

export function comparePassword(rawPassword, passwordHash) {
  return bcrypt.compare(rawPassword, passwordHash)
}
