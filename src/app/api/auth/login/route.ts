import { NextResponse } from 'next/server';
import { z } from 'zod';
import { prisma } from '@/lib/db/prisma';
import { verifyPassword } from '@/lib/auth/argon2';
import { createSession, buildSessionCookie } from '@/lib/auth/session';
import { verifyCsrfOrigin, AuthError } from '@/lib/auth/rbacGuard';
import {
  checkAuthRateLimit,
  recordAuthFailure,
  clearAuthFailure,
  getClientIp,
} from '@/lib/auth/rateLimit';
import { UserStatus } from '@prisma/client';

const loginSchema = z.object({
  identifier: z.string().min(1, 'Username atau email wajib diisi').max(255),
  password: z.string().min(1, 'Password wajib diisi').max(128),
});

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

    const body = await req.json();
    const parsed = loginSchema.safeParse(body);

    if (!parsed.success) {
      return NextResponse.json(
        {
          success: false,
          error: 'Parameter tidak valid',
          details: parsed.error.format(),
        },
        { status: 400 }
      );
    }

    const { identifier, password } = parsed.data;
    const clientIp = getClientIp(req);
    const userAgent = req.headers.get('user-agent') || undefined;

    // 1. Multi-dimensional Rate Limit Check (Per-Account & Per-IP)
    const rateLimit = await checkAuthRateLimit(identifier, clientIp);
    if (!rateLimit.allowed) {
      return NextResponse.json(
        {
          success: false,
          error: rateLimit.reason || 'Terlalu banyak percobaan login.',
          retryAfterSeconds: rateLimit.retryAfterSeconds,
        },
        {
          status: 429,
          headers: {
            'Retry-After': String(rateLimit.retryAfterSeconds || 60),
          },
        }
      );
    }

    // 2. Query user by username or email
    const user = await prisma.user.findFirst({
      where: {
        OR: [{ username: identifier }, { email: identifier }],
      },
      include: {
        roles: {
          include: {
            role: {
              include: {
                permissions: {
                  include: {
                    permission: true,
                  },
                },
              },
            },
          },
        },
      },
    });

    // 3. User not found (generic message to avoid enumeration)
    if (!user) {
      await recordAuthFailure(identifier, clientIp);
      return NextResponse.json(
        {
          success: false,
          error: 'Username atau password salah.',
        },
        { status: 401 }
      );
    }

    // 4. Verify password with Argon2id (performed before holding DB locks)
    const isPasswordValid = await verifyPassword(user.passwordHash, password);
    if (!isPasswordValid) {
      await recordAuthFailure(identifier, clientIp);

      await prisma.auditLog.create({
        data: {
          userId: user.id,
          action: 'USER_LOGIN_FAILED_PASSWORD',
          resourceType: 'User',
          resourceId: user.id,
          ipAddress: clientIp,
          userAgent,
          metadata: { identifier, message: 'Password salah' },
        },
      });

      return NextResponse.json(
        {
          success: false,
          error: 'Username atau password salah.',
        },
        { status: 401 }
      );
    }

    // 5. Authoritative Transaction: Row Lock, status check, session creation, lastLogin update
    // Strictly serialized: validate credentials -> tx -> lock authoritative user -> create session -> COMMIT
    const { sessionToken, expiresAt } = await prisma.$transaction(async (tx) => {
      // Model B Parent User Row Lock
      const lockedRows: any[] = await tx.$queryRawUnsafe(
        `SELECT id, status FROM users WHERE id = $1::uuid FOR UPDATE`,
        user.id
      );

      if (!lockedRows || lockedRows.length === 0) {
        throw new AuthError('User tidak ditemukan.', 401);
      }

      const locked = lockedRows[0];
      if (locked.status !== UserStatus.ACTIVE) {
        if (locked.status === UserStatus.PENDING) {
          throw new AuthError('Akun Anda masih dalam status PENDING dan menunggu persetujuan Administrator.', 401);
        }
        if (locked.status === UserStatus.SUSPENDED) {
          throw new AuthError('Akun Anda telah ditangguhkan (SUSPENDED). Silakan hubungi Administrator.', 401);
        }
        if (locked.status === UserStatus.DISABLED) {
          throw new AuthError('Akun Anda telah dinonaktifkan (DISABLED). Akses tidak diizinkan.', 401);
        }
        throw new AuthError('Status akun tidak valid untuk login.', 401);
      }

      // Create session inside tx (strictly confined to tx)
      const sessionResult = await createSession(user.id, clientIp, userAgent, tx);

      // Update lastLoginAt inside tx
      await tx.user.update({
        where: { id: user.id },
        data: { lastLoginAt: new Date() },
      });

      return sessionResult;
    });

    // 6. Login Success: Clear rate limit counter
    await clearAuthFailure(identifier, clientIp);

    // 7. Record success in Audit Log (after transaction commit)
    await prisma.auditLog.create({
      data: {
        userId: user.id,
        action: 'USER_LOGIN_SUCCESS',
        resourceType: 'User',
        resourceId: user.id,
        ipAddress: clientIp,
        userAgent,
      },
    });

    // Extract roles and permissions
    const roles = user.roles.map((r) => r.role.name);
    const permsSet = new Set<string>();
    user.roles.forEach((r) => {
      r.role.permissions.forEach((p) => permsSet.add(p.permission.code));
    });

    const cookieHeader = buildSessionCookie(sessionToken, expiresAt);

    const response = NextResponse.json(
      {
        success: true,
        message: 'Login berhasil.',
        user: {
          id: user.id,
          username: user.username,
          email: user.email,
          fullName: user.fullName,
          status: user.status,
          isSystemRoot: user.isSystemRoot,
          roles,
          permissions: Array.from(permsSet),
        },
      },
      { status: 200 }
    );

    // Set HttpOnly Secure Cookie
    response.headers.set('Set-Cookie', cookieHeader);
    return response;
  } catch (err: any) {
    if (err instanceof AuthError) {
      return NextResponse.json(
        {
          success: false,
          error: err.message,
        },
        { status: err.statusCode }
      );
    }
    console.error('[API_LOGIN_ERROR]', err);
    return NextResponse.json(
      {
        success: false,
        error: 'Terjadi kesalahan internal pada proses login.',
      },
      { status: 500 }
    );
  }
}
