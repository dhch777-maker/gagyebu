'use client'

import { useState } from 'react'

interface StudentRow {
  id: string
  name: string
  accountId: string
  artworkCount: number
}

const DEMO_STUDENTS: StudentRow[] = [
  { id: '1', name: '최동해', accountId: 'choi_donghae@bunt.kr', artworkCount: 9 },
  { id: '2', name: '이서연', accountId: 'lee_seoyeon@bunt.kr', artworkCount: 6 },
  { id: '3', name: '박지호', accountId: 'park_jiho@bunt.kr', artworkCount: 4 },
]

const inputStyle: React.CSSProperties = {
  width: '100%', background: '#0a0a0a',
  border: 'none', borderBottom: '1px solid #333',
  padding: '8px 0', fontSize: 12, color: '#e0d4c0',
  fontFamily: 'Georgia, serif', outline: 'none',
}

const labelStyle: React.CSSProperties = {
  fontSize: 9, letterSpacing: 2, color: '#444',
  textTransform: 'uppercase', display: 'block', marginBottom: 6,
}

export default function AdminDashboard() {
  const [students, setStudents] = useState<StudentRow[]>(DEMO_STUDENTS)
  const [showAddForm, setShowAddForm] = useState(false)
  const [newName, setNewName] = useState('')
  const [newEmail, setNewEmail] = useState('')
  const [newPassword, setNewPassword] = useState('')
  const [uploadStudentId, setUploadStudentId] = useState<string | null>(null)
  const [uploadTitle, setUploadTitle] = useState('')

  function handleAddStudent(e: React.FormEvent) {
    e.preventDefault()
    setStudents(prev => [...prev, {
      id: Date.now().toString(),
      name: newName,
      accountId: newEmail,
      artworkCount: 0,
    }])
    setNewName(''); setNewEmail(''); setNewPassword('')
    setShowAddForm(false)
  }

  return (
    <div>
      <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 20 }}>
        <button onClick={() => setShowAddForm(v => !v)} style={{
          background: '#d4a853', color: '#0a0a0a', border: 'none',
          padding: '8px 16px', fontSize: 10, letterSpacing: 2,
          textTransform: 'uppercase', cursor: 'pointer', fontFamily: 'Georgia, serif',
        }}>+ 학생 추가</button>
      </div>

      {showAddForm && (
        <form onSubmit={handleAddStudent} style={{
          background: '#111', border: '1px solid #1e1e1e',
          padding: 20, marginBottom: 20,
        }}>
          <p style={{ fontSize: 10, letterSpacing: 2, color: '#d4a853', textTransform: 'uppercase', marginBottom: 16 }}>
            신규 학생 등록
          </p>
          <div style={{ marginBottom: 14 }}>
            <label style={labelStyle}>학생 이름</label>
            <input type="text" value={newName} onChange={e => setNewName(e.target.value)}
              placeholder="김민준" required style={inputStyle} />
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={labelStyle}>학부모 계정 (이메일)</label>
            <input type="email" value={newEmail} onChange={e => setNewEmail(e.target.value)}
              placeholder="parent@email.com" required style={inputStyle} />
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={labelStyle}>초기 비밀번호</label>
            <input type="password" value={newPassword} onChange={e => setNewPassword(e.target.value)}
              placeholder="••••••••" required style={inputStyle} />
          </div>
          <div style={{ display: 'flex', gap: 8, marginTop: 16 }}>
            <button type="submit" style={{
              background: '#d4a853', color: '#0a0a0a', border: 'none',
              padding: '8px 20px', fontSize: 10, letterSpacing: 2,
              textTransform: 'uppercase', cursor: 'pointer', fontFamily: 'Georgia, serif',
            }}>등록</button>
            <button type="button" onClick={() => setShowAddForm(false)} style={{
              background: 'none', color: '#999', border: '1px solid #444',
              padding: '8px 20px', fontSize: 10, letterSpacing: 2,
              textTransform: 'uppercase', cursor: 'pointer', fontFamily: 'Georgia, serif',
            }}>취소</button>
          </div>
        </form>
      )}

      <div style={{ border: '1px solid #1e1e1e' }}>
        {students.map((student, i) => (
          <div key={student.id}>
            <div style={{
              display: 'flex', alignItems: 'center', gap: 16,
              padding: '14px 20px',
              borderBottom: i < students.length - 1 || uploadStudentId === student.id
                ? '1px solid #161616' : 'none',
            }}>
              <div style={{
                width: 36, height: 36, background: '#1a1a1a',
                border: '1px solid #2a2a2a', display: 'flex',
                alignItems: 'center', justifyContent: 'center',
                fontSize: 16, flexShrink: 0,
              }}>👤</div>
              <div style={{ flex: 1 }}>
                <div style={{ fontSize: 13, color: '#e0d4c0', letterSpacing: 1 }}>{student.name}</div>
                <div style={{ fontSize: 10, color: '#888', letterSpacing: 1, marginTop: 2 }}>
                  {student.artworkCount}점 · {student.accountId}
                </div>
              </div>
              <button
                onClick={() => setUploadStudentId(uploadStudentId === student.id ? null : student.id)}
                style={{
                  background: 'none', color: '#d4a85380',
                  border: '1px solid #d4a85330', padding: '5px 12px',
                  fontSize: 9, letterSpacing: 1, textTransform: 'uppercase',
                  cursor: 'pointer', fontFamily: 'Georgia, serif',
                }}>사진 올리기</button>
            </div>

            {uploadStudentId === student.id && (
              <div style={{ background: '#0d0d0d', padding: '16px 20px', borderBottom: '1px solid #161616' }}>
                <p style={{ fontSize: 9, letterSpacing: 2, color: '#d4a853', textTransform: 'uppercase', marginBottom: 12 }}>
                  {student.name} · 작품 업로드
                </p>
                <div style={{ marginBottom: 12 }}>
                  <label style={labelStyle}>작품명</label>
                  <input type="text" value={uploadTitle} onChange={e => setUploadTitle(e.target.value)}
                    placeholder="작품 제목" style={inputStyle} />
                </div>
                <div style={{ marginBottom: 16 }}>
                  <label style={labelStyle}>사진 파일 (다중 선택 가능)</label>
                  <input type="file" accept="image/*" multiple
                    style={{ fontSize: 11, color: '#555' }} />
                </div>
                <button style={{
                  background: '#d4a853', color: '#0a0a0a', border: 'none',
                  padding: '8px 20px', fontSize: 10, letterSpacing: 2,
                  textTransform: 'uppercase', cursor: 'pointer', fontFamily: 'Georgia, serif',
                }}>업로드</button>
              </div>
            )}
          </div>
        ))}
      </div>
    </div>
  )
}
