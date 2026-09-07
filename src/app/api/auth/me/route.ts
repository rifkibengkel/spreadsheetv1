import { NextResponse } from 'next/server';
import { authenticateRequest, unauthorizedResponse } from '@/lib/auth/rbacGuard';

export async function GET(req: Request) {
  try {
    const user = await authenticateRequest(req);
    if (!user) {
      return unauthorizedResponse('Sesi tidak ditemukan atau telah kedaluwarsa.');
    }

    return NextResponse.json(
      {
        success: true,
        user: {
          id: user.id,
          username: user.username,
          email: user.email,
          fullName: user.fullName,
          status: user.status,
          isSystemRoot: user.isSystemRoot,
          roles: user.roles,
          permissions: user.permissions,
        },
      },
      { status: 200 }
    );
  } catch (err: any) {
    console.error('[API_ME_ERROR]', err);
    return NextResponse.json(
      { success: false, error: 'Gagal mengambil data profil sesi.' },
      { status: 500 }
    );
  }
}
