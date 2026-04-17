import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { pool } from './pool.js'

const currentFile = fileURLToPath(import.meta.url)
const currentDir = path.dirname(currentFile)
const schemaPath = path.join(currentDir, 'schema.sql')

async function runMigrations() {
  const sql = await readFile(schemaPath, 'utf8')
  await pool.query(sql)
  await pool.end()
  console.log('Database migration completed.')
}

runMigrations().catch(async (error) => {
  console.error('Migration failed:', error.message)
  await pool.end()
  process.exit(1)
})
