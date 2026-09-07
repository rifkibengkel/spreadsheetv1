import { NextResponse } from 'next/server';
import { requireAuth, requirePermission, requireRootAdmin } from '@/lib/auth/rbacGuard';

export async function GET(req: Request) {
  try {
    const url = new URL(req.url);
    const requiredPerm = url.searchParams.get('perm');
    const isRootRequired = url.searchParams.get('requireRoot') === 'true';

    let user;

    if (isRootRequired) {
      user = await requireRootAdmin(req);
    } else if (requiredPerm) {
      user = await requirePermission(req, requiredPerm);
    } else {
      user = await requireAuth(req);
    }

    return NextResponse.json(
      {
        success: true,
        message: 'Akses diizinkan.',
        checkedPermission: requiredPerm || (isRootRequired ? 'ROOT_ONLY' : 'AUTHENTICATED_ONLY'),
        user: {
          id: user.id,
          username: user.username,
          isSystemRoot: user.isSystemRoot,
          roles: user.roles,
        },
      },
      { status: 200 }
    );
  } catch (err: any) {
    const statusCode = err.statusCode || 500;
    return NextResponse.json(
      {
        success: false,
        error: err.message || 'Akses ditolak.',
      },
      { status: statusCode }
    );
  }
}
