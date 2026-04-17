import { Router } from 'express'
import { env } from '../config/env.js'

export const healthRouter = Router()

healthRouter.get('/', (_req, res) => {
  res.json({
    status: 'ok',
    service: 'utalk-api',
    environment: env.NODE_ENV,
    publicUrls: {
      app: env.APP_PUBLIC_URL,
      web: env.WEB_PUBLIC_URL,
      api: env.API_PUBLIC_URL,
      ws: env.WS_PUBLIC_URL,
    },
  })
})
