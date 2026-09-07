import { NextResponse } from 'next/server';
import {
  extractSessionToken,
  validateSessionToken,
  revokeSession,
  buildLogoutCookie,
} from '@/lib/auth/session';
import { verifyCsrfOrigin } from '@/lib/auth/rbacGuard';
import { prisma } from '@/lib/db/prisma';
import { getClientIp } from '@/lib/auth/rateLimit';

export async function POST(req: Request) {
  try {
    // 0. CSRF Origin Verification
    if (!verifyCsrfOrigin(req)) {
      return NextResponse.json(
        {
          success: false,
          error: 'Akses ditolak: Verifikasi CSRF origin gagal.',
        },
        { status: 403 }
      );
    }

    const token = extractSessionToken(req);
    const clientIp = getClientIp(req);
    const userAgent = req.headers.get('user-agent') || undefined;

    if (token) {
      // Optional: fetch user for audit log before revocation
      const user = await validateSessionToken(token);

      // Authoritative revocation from PostgreSQL
      await revokeSession(token);

      if (user) {
        await prisma.auditLog.create({
          data: {
            userId: user.id,
            action: 'USER_LOGOUT',
            resourceType: 'User',
            resourceId: user.id,
            ipAddress: clientIp,
            userAgent,
          },
        });
      }
    }

    const response = NextResponse.json(
      {
        success: true,
        message: 'Logout berhasil.',
      },
      { status: 200 }
    );

    // Clear session cookie
    response.headers.set('Set-Cookie', buildLogoutCookie());
    return response;
  } catch (err: any) {
    console.error('[API_LOGOUT_ERROR]', err);
    return NextResponse.json(
      {
        success: false,
        error: 'Gagal memproses logout.',
      },
      { status: 500 }
    );
  }
}

