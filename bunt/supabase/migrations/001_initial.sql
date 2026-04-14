-- ============================================================
-- BUNT Kunstschule Magok - Initial Schema Migration
-- 001_initial.sql
-- ============================================================

-- ============================================================
-- 1. TABLES
-- ============================================================

-- students: 학생 정보
CREATE TABLE IF NOT EXISTS students (
  id         uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  name       text NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now()
);

-- parent_accounts: 학부모 계정 (Supabase Auth 연동)
CREATE TABLE IF NOT EXISTS parent_accounts (
  id           uuid PRIMARY KEY REFERENCES auth.users(id) ON DELETE CASCADE,
  student_id   uuid NOT NULL REFERENCES students(id) ON DELETE CASCADE,
  display_name text NOT NULL
);

-- artworks: 작품 정보
CREATE TABLE IF NOT EXISTS artworks (
  id          uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  student_id  uuid NOT NULL REFERENCES students(id) ON DELETE CASCADE,
  title       text NOT NULL,
  image_path  text NOT NULL,
  created_at  timestamptz NOT NULL DEFAULT now()
);

-- ============================================================
-- 2. ROW LEVEL SECURITY (RLS)
-- ============================================================

-- 2-1. students 테이블 RLS
ALTER TABLE students ENABLE ROW LEVEL SECURITY;

-- 학부모는 자신의 student_id와 일치하는 학생만 조회 가능
CREATE POLICY "parents can view own student"
  ON students FOR SELECT
  USING (
    id IN (
      SELECT student_id
      FROM parent_accounts
      WHERE id = auth.uid()
    )
  );

-- INSERT: service_role만 허용 (정책 없음 = 일반 사용자 거부)
-- UPDATE: service_role만 허용
-- DELETE: service_role만 허용
-- (RLS가 활성화된 상태에서 매칭 정책이 없으면 자동 거부됨)

-- 2-2. parent_accounts 테이블 RLS
ALTER TABLE parent_accounts ENABLE ROW LEVEL SECURITY;

-- 본인 계정만 조회 가능
CREATE POLICY "users can view own parent account"
  ON parent_accounts FOR SELECT
  USING (id = auth.uid());

-- INSERT/UPDATE/DELETE: service_role만 허용 (정책 없음 = 일반 사용자 거부)

-- 2-3. artworks 테이블 RLS
ALTER TABLE artworks ENABLE ROW LEVEL SECURITY;

-- 학부모는 자신의 학생 작품만 조회 가능
CREATE POLICY "parents can view own student artworks"
  ON artworks FOR SELECT
  USING (
    student_id IN (
      SELECT student_id
      FROM parent_accounts
      WHERE id = auth.uid()
    )
  );

-- INSERT/UPDATE/DELETE: service_role만 허용 (정책 없음 = 일반 사용자 거부)

-- ============================================================
-- 3. STORAGE
-- ============================================================

-- artworks 버킷 생성 (비공개)
INSERT INTO storage.buckets (id, name, public)
VALUES ('artworks', 'artworks', false)
ON CONFLICT (id) DO NOTHING;

-- 학부모가 자신의 학생 작품 이미지만 다운로드 가능
-- Storage 경로: artworks/{student_id}/{filename}
CREATE POLICY "parents can view own student artworks"
  ON storage.objects FOR SELECT
  USING (
    bucket_id = 'artworks'
    AND (storage.foldername(name))[1] IN (
      SELECT student_id::text
      FROM parent_accounts
      WHERE id = auth.uid()
    )
  );

-- service_role은 모든 스토리지 작업 가능 (이미지 업로드 등)
CREATE POLICY "service role full access"
  ON storage.objects
  USING (auth.role() = 'service_role')
  WITH CHECK (auth.role() = 'service_role');
