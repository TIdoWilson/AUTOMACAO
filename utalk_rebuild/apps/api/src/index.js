import { buildApp } from './app.js'
import { env } from './config/env.js'
import { pool } from './db/pool.js'

const app = buildApp()

async function bootstrap() {
  await pool.query('SELECT 1')
  app.listen(env.PORT, () => {
    console.log(`UTalk API running on port ${env.PORT}`)
    console.log(`Public API URL: ${env.API_PUBLIC_URL}`)
  })
}

bootstrap().catch(async (error) => {
  console.error('Failed to start API:', error.message)
  await pool.end()
  process.exit(1)
})
