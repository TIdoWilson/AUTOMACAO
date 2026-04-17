import cookieParser from 'cookie-parser'
import cors from 'cors'
import express from 'express'
import { env } from './config/env.js'
import { authRouter } from './routes/auth.js'
import { contactsRouter } from './routes/contacts.js'
import { conversationsRouter } from './routes/conversations.js'
import { dashboardRouter } from './routes/dashboard.js'
import { departmentsRouter } from './routes/departments.js'
import { healthRouter } from './routes/health.js'
import { settingsRouter } from './routes/settings.js'
import { usersRouter } from './routes/users.js'

export function buildApp() {
  const app = express()

  app.use(
    cors({
      origin: [env.WEB_PUBLIC_URL, env.APP_PUBLIC_URL],
      credentials: true,
    }),
  )
  app.use(express.json({ limit: '2mb' }))
  app.use(cookieParser())

  app.use('/health', healthRouter)
  app.use('/auth', authRouter)
  app.use('/departments', departmentsRouter)
  app.use('/users', usersRouter)
  app.use('/settings', settingsRouter)
  app.use('/contacts', contactsRouter)
  app.use('/conversations', conversationsRouter)
  app.use('/dashboard', dashboardRouter)

  app.use((req, res) => {
    res.status(404).json({ error: `Route not found: ${req.method} ${req.path}` })
  })

  app.use((error, _req, res, _next) => {
    const status = Number.isInteger(error?.statusCode) ? error.statusCode : 500
    const message = error?.message ?? 'Unexpected error.'
    if (status >= 500) {
      console.error(error)
    }
    res.status(status).json({ error: message })
  })

  return app
}
