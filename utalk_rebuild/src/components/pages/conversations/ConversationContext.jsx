import { createContext, useContext } from 'react'

const ConversationContext = createContext(null)

export function ConversationProvider({ value, children }) {
  return <ConversationContext.Provider value={value}>{children}</ConversationContext.Provider>
}

export function useConversationContext() {
  const context = useContext(ConversationContext)
  if (!context) {
    throw new Error('useConversationContext deve ser usado dentro de ConversationProvider')
  }
  return context
}
