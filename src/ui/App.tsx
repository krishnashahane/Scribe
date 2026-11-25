import { useEffect, useState } from 'react'
import { api, type User } from '../api'
import { ToastProvider } from './kit'
import { Login } from './Login'
import { Home } from './Home'
import { EditorScreen } from './EditorScreen'

export type Theme = 'light' | 'dark'

function useRoute() {
  const [hash, setHash] = useState(location.hash)
  useEffect(() => {
    const on = () => setHash(location.hash)
    window.addEventListener('hashchange', on)
    return () => window.removeEventListener('hashchange', on)
  }, [])
  const m = /^#\/doc\/([A-Za-z0-9_-]+)/.exec(hash)
  return m ? { name: 'doc' as const, id: m[1] } : { name: 'home' as const }
}

export function App() {
  const [user, setUser] = useState<User | null | undefined>(undefined)
  const [theme, setTheme] = useState<Theme>(() => (localStorage.getItem('scribe-theme') as Theme) || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'))
  const route = useRoute()

  useEffect(() => {
    document.documentElement.dataset.theme = theme
    localStorage.setItem('scribe-theme', theme)
  }, [theme])

  useEffect(() => {
    api<{ user: User | null }>('/api/auth/me')
      .then((r) => setUser(r.user))
      .catch(() => setUser(null))
  }, [])

  const toggleTheme = () => setTheme((t) => (t === 'dark' ? 'light' : 'dark'))
  const signOut = async () => {
    await api('/api/auth/logout', { method: 'POST' }).catch(() => {})
    setUser(null)
    location.hash = ''
  }

  return (
    <ToastProvider>
      {user === undefined ? (
        <div className="sc-splash" aria-busy="true">
          <span className="sc-logo">SCRIBE</span>
        </div>
      ) : user === null ? (
        <Login onAuth={setUser} />
      ) : route.name === 'doc' ? (
        <EditorScreen key={route.id} docId={route.id} user={user} theme={theme} onToggleTheme={toggleTheme} />
      ) : (
        <Home user={user} onSignOut={signOut} theme={theme} onToggleTheme={toggleTheme} />
      )}
    </ToastProvider>
  )
}
