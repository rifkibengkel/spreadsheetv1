import { NextResponse } from 'next/server';
import { z } from 'zod';
import { requirePermission, verifyCsrfOrigin, AuthError } from '@/lib/auth/rbacGuard';
import { extractSessionToken, hashSessionToken } from '@/lib/auth/session';
import { updateUser, deleteUser, AuthForbiddenError, AuthNotFoundError } from '@/lib/auth/userService';
import { UserStatus } from '@prisma/client';

const updateUserSchema = z.object({
  fullName: z.string().min(1).max(100).optional(),
  status: z.nativeEnum(UserStatus).optional(),
  isSystemRoot: z.boolean().optional(),
  roleNames: z.array(z.string()).optional(),
});

export async function PATCH(
  req: Request,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    // 0. CSRF Origin Verification
    if (!verifyCsrfOrigin(req)) {
      return NextResponse.json(
        { success: false, error: 'Akses ditolak: Verifikasi CSRF origin gagal.' },
        { status: 403 }
      );
    }

    // 1. Early request screening & fast-fail
    await requirePermission(req, 'USER_MANAGE');

    // 2. Extract session token for authoritative transactional verification
    const token = extractSessionToken(req);
    if (!token) {
      throw new AuthError('Otentikasi diperlukan.', 401);
    }
    const sessionTokenHash = hashSessionToken(token);

    const { id: targetUserId } = await params;
    const body = await req.json();
    const parsed = updateUserSchema.safeParse(body);

    if (!parsed.success) {
      return NextResponse.json(
        { success: false, error: 'Payload tidak valid.', details: parsed.error.format() },
        { status: 400 }
      );
    }

    // 3. Delegate to transactional service with authoritative DB verification
    await updateUser({
      sessionTokenHash,
      targetUserId,
      payload: parsed.data,
    });

    return NextResponse.json(
      { success: true, message: 'User berhasil diperbarui.' },
      { status: 200 }
    );
  } catch (err: any) {
    if (err instanceof AuthError) {
      return NextResponse.json(
        { success: false, error: err.message },
        { status: err.statusCode }
      );
    }
    if (err instanceof AuthForbiddenError || err instanceof AuthNotFoundError) {
      return NextResponse.json(
        { success: false, error: err.message },
        { status: err.statusCode }
      );
    }
    console.error('[ADMIN_USER_PATCH_ERROR]', err);
    return NextResponse.json(
      { success: false, error: 'Terjadi kesalahan server internal.' },
      { status: 500 }
    );
  }
}

export async function DELETE(
  req: Request,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    // 0. CSRF Origin Verification
    if (!verifyCsrfOrigin(req)) {
      return NextResponse.json(
        { success: false, error: 'Akses ditolak: Verifikasi CSRF origin gagal.' },
        { status: 403 }
      );
    }

    // 1. Early request screening & fast-fail
    await requirePermission(req, 'USER_MANAGE');

    // 2. Extract session token for authoritative transactional verification
    const token = extractSessionToken(req);
    if (!token) {
      throw new AuthError('Otentikasi diperlukan.', 401);
    }
    const sessionTokenHash = hashSessionToken(token);

    const { id: targetUserId } = await params;

    // 3. Delegate to transactional service with authoritative DB verification
    await deleteUser({
      sessionTokenHash,
      targetUserId,
    });

    return NextResponse.json(
      { success: true, message: 'User berhasil dihapus.' },
      { status: 200 }
    );
  } catch (err: any) {
    if (err instanceof AuthError) {
      return NextResponse.json(
        { success: false, error: err.message },
        { status: err.statusCode }
      );
    }
    if (err instanceof AuthForbiddenError || err instanceof AuthNotFoundError) {
      return NextResponse.json(
        { success: false, error: err.message },
        { status: err.statusCode }
      );
    }
    console.error('[ADMIN_USER_DELETE_ERROR]', err);
    return NextResponse.json(
      { success: false, error: 'Terjadi kesalahan server internal.' },
      { status: 500 }
    );
  }
}

