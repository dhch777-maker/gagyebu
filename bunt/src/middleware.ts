import { createServerClient } from '@supabase/ssr'
import { NextResponse, type NextRequest } from 'next/server'

export async function middleware(request: NextRequest) {
  let supabaseResponse = NextResponse.next({ request })

  const supabase = createServerClient(
    process.env.NEXT_PUBLIC_SUPABASE_URL!,
    process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY!,
    {
      cookies: {
        getAll() { return request.cookies.getAll() },
        setAll(cookiesToSet) {
          cookiesToSet.forEach(({ name, value }) => request.cookies.set(name, value))
          supabaseResponse = NextResponse.next({ request })
          cookiesToSet.forEach(({ name, value, options }) =>
            supabaseResponse.cookies.set(name, value, options)
          )
        },
      },
    }
  )

  const { data: { user } } = await supabase.auth.getUser()

  const path = request.nextUrl.pathname

  // 미인증 사용자가 /gallery 또는 /admin 접근 시 → / 리다이렉트
  if (!user && (path.startsWith('/gallery') || path.startsWith('/admin'))) {
    return NextResponse.redirect(new URL('/', request.url))
  }

  // /admin 접근 시 관리자 역할 확인
  if (user && path.startsWith('/admin') && user.app_metadata?.role !== 'admin') {
    return NextResponse.redirect(new URL('/gallery', request.url))
  }

  // 로그인 상태에서 / 접근 시 → /gallery 리다이렉트
  if (user && path === '/') {
    return NextResponse.redirect(new URL('/gallery', request.url))
  }

  return supabaseResponse
}

export const config = {
  matcher: ['/', '/gallery', '/gallery/:path*', '/admin', '/admin/:path*'],
}
