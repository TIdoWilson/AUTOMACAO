import { config as loadEnv } from 'dotenv'
import { z } from 'zod'

loadEnv()

const schema = z.object({
  NODE_ENV: z.enum(['development', 'test', 'production']).default('development'),
  PORT: z.coerce.number().int().positive().default(4000),
  DATABASE_URL: z.string().min(1, 'DATABASE_URL is required'),
  DB_SSL: z.string().optional().default('false'),
  APP_PUBLIC_URL: z.string().url(),
  WEB_PUBLIC_URL: z.string().url(),
  API_PUBLIC_URL: z.string().url(),
  WS_PUBLIC_URL: z.string().url(),
  JWT_ACCESS_SECRET: z.string().min(16, 'JWT_ACCESS_SECRET must have at least 16 chars'),
  ACCESS_TOKEN_TTL: z.string().default('15m'),
  REFRESH_TOKEN_TTL_DAYS: z.coerce.number().int().positive().default(7),
  COOKIE_SECURE: z.string().optional().default('false'),
  INTERNAL_API_TOKEN: z.string().min(12),
})

const parsed = schema.safeParse(process.env)

if (!parsed.success) {
  const errors = parsed.error.issues.map((issue) => `${issue.path.join('.')}: ${issue.message}`).join('\n')
  throw new Error(`Invalid environment variables:\n${errors}`)
}

function assertNoLocalhost(name, value) {
  const url = new URL(value)
  const host = url.hostname.toLowerCase()
  if (host === 'localhost' || host === '127.0.0.1') {
    throw new Error(`${name} cannot use localhost or 127.0.0.1. Configure machine IP instead.`)
  }
}

assertNoLocalhost('APP_PUBLIC_URL', parsed.data.APP_PUBLIC_URL)
assertNoLocalhost('WEB_PUBLIC_URL', parsed.data.WEB_PUBLIC_URL)
assertNoLocalhost('API_PUBLIC_URL', parsed.data.API_PUBLIC_URL)
assertNoLocalhost('WS_PUBLIC_URL', parsed.data.WS_PUBLIC_URL)

export const env = {
  ...parsed.data,
  DB_SSL: parsed.data.DB_SSL === 'true',
  COOKIE_SECURE: parsed.data.COOKIE_SECURE === 'true',
}
