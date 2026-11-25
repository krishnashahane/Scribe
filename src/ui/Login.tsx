import { useState } from 'react'
import { api, type User } from '../api'

export function Login({ onAuth }: { onAuth: (u: User) => void }) {
  const [mode, setMode] = useState<'login' | 'register'>('login')
  const [email, setEmail] = useState('')
  const [name, setName] = useState('')
  const [password, setPassword] = useState('')
  const [error, setError] = useState('')
  const [busy, setBusy] = useState(false)

  const submit = async (e: React.FormEvent) => {
    e.preventDefault()
    setBusy(true)
    setError('')
    try {
      const r = await api<{ user: User }>(`/api/auth/${mode}`, { method: 'POST', body: { email, password, name } })
      onAuth(r.user)
    } catch (err: any) {
      setError(err.message)
    } finally {
      setBusy(false)
    }
  }

  return (
    <main className="sc-auth">
      <section className="sc-auth-card">
        <div className="sc-auth-brand">
          <span className="sc-logo">SCRIBE</span>
          <p>A word processor that respects your documents.</p>
        </div>
        <form onSubmit={submit} aria-label={mode === 'login' ? 'Sign in' : 'Create account'}>
          <h1>{mode === 'login' ? 'Sign in' : 'Create your account'}</h1>
          {mode === 'register' && (
            <label>
              Name
              <input value={name} onChange={(e) => setName(e.target.value)} autoComplete="name" maxLength={100} />
            </label>
          )}
          <label>
            Email
            <input type="email" required value={email} onChange={(e) => setEmail(e.target.value)} autoComplete="email" />
          </label>
          <label>
            Password
            <input type="password" required minLength={8} value={password} onChange={(e) => setPassword(e.target.value)} autoComplete={mode === 'login' ? 'current-password' : 'new-password'} />
          </label>
          {error && (
            <p className="sc-error" role="alert">
              {error}
            </p>
          )}
          <button className="sc-primary" type="submit" disabled={busy}>
            {busy ? 'Please wait…' : mode === 'login' ? 'Sign in' : 'Create account'}
          </button>
          <p className="sc-auth-switch">
            {mode === 'login' ? 'New to SCRIBE?' : 'Already have an account?'}{' '}
            <button type="button" className="sc-link" onClick={() => { setMode(mode === 'login' ? 'register' : 'login'); setError('') }}>
              {mode === 'login' ? 'Create an account' : 'Sign in'}
            </button>
          </p>
        </form>
      </section>
    </main>
  )
}
