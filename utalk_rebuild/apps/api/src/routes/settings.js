import { Router } from 'express'
import { requireAuth, requireRole } from '../middleware/auth.js'

export const settingsRouter = Router()

const SETTINGS_MODULES = [
  'channels',
  'financial',
  'departments',
  'ai-agents',
  'knowledge-bases',
  'attendants',
  'tags',
  'chatbots',
  'quick-replies',
  'whatsapp-templates',
]

settingsRouter.get('/modules', requireAuth, requireRole('admin'), (_req, res) => {
  res.json({ modules: SETTINGS_MODULES })
})
