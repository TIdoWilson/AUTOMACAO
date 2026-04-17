const ACCESS_TOKEN_KEY = 'utalk_access_token'
const USER_KEY = 'utalk_user'

function defaultApiBaseUrl() {
  if (typeof window === 'undefined') {
    return 'http://0.0.0.0:4000'
  }
  return `${window.location.protocol}//${window.location.hostname}:4000`
}

export const API_BASE_URL = (import.meta.env.VITE_API_BASE_URL || '').trim() || defaultApiBaseUrl()
export const INTERNAL_API_TOKEN = (import.meta.env.VITE_INTERNAL_API_TOKEN || '').trim()

export function getStoredAccessToken() {
  return localStorage.getItem(ACCESS_TOKEN_KEY)
}

export function getStoredUser() {
  try {
    const raw = localStorage.getItem(USER_KEY)
    return raw ? JSON.parse(raw) : null
  } catch {
    return null
  }
}

export function storeSession({ accessToken, user }) {
  localStorage.setItem(ACCESS_TOKEN_KEY, accessToken)
  localStorage.setItem(USER_KEY, JSON.stringify(user))
}

export function clearSession() {
  localStorage.removeItem(ACCESS_TOKEN_KEY)
  localStorage.removeItem(USER_KEY)
}

async function parseResponse(response) {
  const text = await response.text()
  let data = null
  if (text) {
    try {
      data = JSON.parse(text)
    } catch {
      data = { raw: text }
    }
  }
  if (!response.ok) {
    const message = data?.error || `Request failed (${response.status})`
    throw new Error(message)
  }
  return data
}

export async function apiRequest(path, { method = 'GET', token = null, body = null, headers = {} } = {}) {
  const requestHeaders = {
    ...headers,
  }
  if (body !== null && !requestHeaders['Content-Type']) {
    requestHeaders['Content-Type'] = 'application/json'
  }
  if (token) {
    requestHeaders.Authorization = `Bearer ${token}`
  }

  const response = await fetch(`${API_BASE_URL}${path}`, {
    method,
    credentials: 'include',
    headers: requestHeaders,
    body: body !== null ? JSON.stringify(body) : undefined,
  })
  return parseResponse(response)
}

export function loginRequest(payload) {
  return apiRequest('/auth/login', { method: 'POST', body: payload })
}

export function meRequest(token) {
  return apiRequest('/auth/me', { token })
}

export function refreshRequest() {
  return apiRequest('/auth/refresh', { method: 'POST' })
}

export function logoutRequest(token) {
  return apiRequest('/auth/logout', { method: 'POST', token })
}

export function listDepartmentsRequest(token) {
  return apiRequest('/departments', { token })
}

export function listContactsRequest(token, search = '') {
  const query = search ? `?search=${encodeURIComponent(search)}` : ''
  return apiRequest(`/contacts${query}`, { token })
}

export function listConversationsRequest(token) {
  return apiRequest('/conversations', { token })
}

export function listConversationMessagesRequest(token, conversationId) {
  return apiRequest(`/conversations/${conversationId}/messages`, { token })
}

export function assumeConversationRequest(token, conversationId) {
  return apiRequest(`/conversations/${conversationId}/assume`, { method: 'POST', token })
}

export function leaveConversationRequest(token, conversationId) {
  return apiRequest(`/conversations/${conversationId}/leave`, { method: 'POST', token })
}

export function finalizeConversationRequest(token, conversationId) {
  return apiRequest(`/conversations/${conversationId}/finalize`, { method: 'POST', token })
}

export function waitCustomerConversationRequest(token, conversationId) {
  return apiRequest(`/conversations/${conversationId}/waiting-customer`, { method: 'POST', token })
}

export function transferConversationRequest(token, conversationId, departmentId) {
  return apiRequest(`/conversations/${conversationId}/transfer`, {
    method: 'POST',
    token,
    body: { departmentId },
  })
}

export function sendConversationMessageRequest(token, conversationId, body) {
  return apiRequest(`/conversations/${conversationId}/messages`, {
    method: 'POST',
    token,
    body: { body },
  })
}

export function incomingMessageRequest(payload) {
  if (!INTERNAL_API_TOKEN) {
    throw new Error('VITE_INTERNAL_API_TOKEN is required to simulate customer incoming messages.')
  }
  return apiRequest('/conversations/incoming-message', {
    method: 'POST',
    headers: { 'x-internal-token': INTERNAL_API_TOKEN },
    body: payload,
  })
}
