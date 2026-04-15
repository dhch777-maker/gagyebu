import NavBar from '@/components/NavBar'
import AdminDashboard from '@/components/AdminDashboard'

export default function AdminPage() {
  return (
    <div style={{ minHeight: '100vh', background: '#0a0a0a', fontFamily: 'Georgia, serif' }}>
      <NavBar isAdmin />
      <main style={{ maxWidth: 800, margin: '0 auto', padding: '32px 24px' }}>
        <div style={{ marginBottom: 28, paddingBottom: 16, borderBottom: '1px solid #1e1e1e' }}>
          <p style={{ fontSize: 10, letterSpacing: 3, color: '#d4a85360', textTransform: 'uppercase', marginBottom: 6 }}>
            Admin · 관리자
          </p>
          <h1 style={{ fontSize: 22, color: '#e0d4c0', letterSpacing: 2 }}>학생 관리</h1>
        </div>
        <AdminDashboard />
      </main>
    </div>
  )
}
