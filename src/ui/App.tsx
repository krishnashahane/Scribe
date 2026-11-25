import { useEffect, useState } from 'react'
import { api, type User } from '../api'
import { ToastProvider } from './kit'
import { Login } from './Login'
import { Home } from './Home'
import { EditorScreen } from './EditorScreen'
import { createFromTemplate, openDoc } from './docActions'

export type Theme = 'light' | 'dark'

type Route = { name: 'doc'; id: string } | { name: 'documents' } | { name: 'signin' } | { name: 'start' }

function parse(hash: string): Route {
  const m = /^#\/doc\/([A-Za-z0-9_-]+)/.exec(hash)
  if (m) return { name: 'doc', id: m[1] }
  if (hash.startsWith('#/documents')) return { name: 'documents' }
  if (hash.startsWith('#/signin')) return { name: 'signin' }
  return { name: 'start' }
}

function useRoute() {
  const [route, setRoute] = useState(() => parse(location.hash))
  useEffect(() => {
    const on = () => setRoute(parse(location.hash))
    window.addEventListener('hashchange', on)
    return () => window.removeEventListener('hashchange', on)
  }, [])
  return route
}

export function App() {
  const [user, setUser] = useState<User | null | undefined>(undefined)
  const [startError, setStartError] = useState('')
  const [theme, setTheme] = useState<Theme>(() => (localStorage.getItem('scribe-theme') as Theme) || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'))
  const route = useRoute()

  useEffect(() => {
    document.documentElement.dataset.theme = theme
    localStorage.setItem('scribe-theme', theme)
  }, [theme])

  // No sign-in wall: reuse the browser's session, or start a private local one.
  useEffect(() => {
    api<{ user: User | null }>('/api/auth/me')
      .then((r) => r.user ?? api<{ user: User }>('/api/auth/guest', { method: 'POST' }).then((g) => g.user))
      .then(setUser)
      .catch((e) => {
        setStartError(e.message)
        setUser(null)
      })
  }, [])

  // Open straight into the last document, or a fresh blank one.
  useEffect(() => {
    if (!user || route.name !== 'start') return
    const last = localStorage.getItem('scribe-last-doc')
    const fresh = () =>
      createFromTemplate('blank')
        .then(openDoc)
        .catch((e) => setStartError(e.message))
    if (!last) return void fresh()
    api(`/api/docs/${last}`)
      .then(() => openDoc(last))
      .catch(() => {
        localStorage.removeItem('scribe-last-doc')
        fresh()
      })
  }, [user, route.name])

  const toggleTheme = () => setTheme((t) => (t === 'dark' ? 'light' : 'dark'))
  const signOut = async () => {
    await api('/api/auth/logout', { method: 'POST' }).catch(() => {})
    localStorage.removeItem('scribe-last-doc')
    setUser(null)
    location.hash = '#/signin'
  }
  const onAuth = (u: User) => {
    localStorage.removeItem('scribe-last-doc')
    setUser(u)
    location.hash = '#/documents'
  }

  let screen
  if (user === undefined || (user && route.name === 'start' && !startError)) {
    screen = (
      <div className="sc-splash" aria-busy="true">
        <span className="sc-logo">SCRIBE</span>
        <span className="sc-muted">Opening…</span>
      </div>
    )
  } else if (!user || route.name === 'signin') {
    screen = (
      <Login
        onAuth={onAuth}
        notice={startError}
        onContinue={
          user
            ? () => (location.hash = '')
            : () =>
                api<{ user: User }>('/api/auth/guest', { method: 'POST' })
                  .then((g) => {
                    setUser(g.user)
                    location.hash = ''
                  })
                  .catch((e) => setStartError(e.message))
        }
      />
    )
  } else if (route.name === 'doc') {
    screen = <EditorScreen key={route.id} docId={route.id} user={user} onUser={setUser} theme={theme} onToggleTheme={toggleTheme} />
  } else {
    screen = <Home user={user} onUser={setUser} onSignOut={signOut} theme={theme} onToggleTheme={toggleTheme} />
  }
  return <ToastProvider>{screen}</ToastProvider>
}
