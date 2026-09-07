import { NextResponse } from 'next/server';
import { z } from 'zod';
import { prisma } from '@/lib/db/prisma';
import { hashPassword } from '@/lib/auth/argon2';
import { getClientIp } from '@/lib/auth/rateLimit';
import { UserStatus } from '@prisma/client';

const registerSchema = z.object({
  username: z
    .string()
    .min(3, 'Username minimal 3 karakter')
    .max(30, 'Username maksimal 30 karakter')
    .regex(/^[a-zA-Z0-9_-]+$/, 'Username hanya boleh huruf, angka, underscore, dan dash'),
  email: z.string().email('Format email tidak valid').max(255),
  password: z
    .string()
    .min(8, 'Password minimal 8 karakter')
    .max(128, 'Password maksimal 128 karakter'),
  fullName: z
    .string()
    .min(2, 'Nama lengkap minimal 2 karakter')
    .max(100, 'Nama lengkap maksimal 100 karakter'),
});

export async function POST(req: Request) {
  try {
    const body = await req.json();
    const parsed = registerSchema.safeParse(body);

    if (!parsed.success) {
      return NextResponse.json(
        {
          success: false,
          error: 'Validasi gagal',
          details: parsed.error.format(),
        },
        { status: 400 }
      );
    }

    const { username, email, password, fullName } = parsed.data;
    const clientIp = getClientIp(req);
    const userAgent = req.headers.get('user-agent') || undefined;

    // Check duplicate username or email
    const existing = await prisma.user.findFirst({
      where: {
        OR: [{ username }, { email }],
      },
      select: { username: true, email: true },
    });

    if (existing) {
      const field = existing.username === username ? 'Username' : 'Email';
      return NextResponse.json(
        {
          success: false,
          error: `${field} sudah digunakan oleh akun lain.`,
        },
        { status: 409 }
      );
    }

    // Hash password with Argon2id
    const passwordHash = await hashPassword(password);

    // Create user with PENDING status (Admin approval flow)
    const newUser = await prisma.user.create({
      data: {
        username,
        email,
        passwordHash,
        fullName,
        status: UserStatus.PENDING,
        isSystemRoot: false,
      },
    });

    // Assign standard USER role
    const userRole = await prisma.role.findUnique({ where: { name: 'USER' } });
    if (userRole) {
      await prisma.userRole.create({
        data: {
          userId: newUser.id,
          roleId: userRole.id,
        },
      });
    }

    // Log registration in Audit Logs
    await prisma.auditLog.create({
      data: {
        userId: newUser.id,
        action: 'USER_REGISTERED_PENDING',
        resourceType: 'User',
        resourceId: newUser.id,
        ipAddress: clientIp,
        userAgent,
        metadata: {
          username,
          email,
          status: 'PENDING',
          message: 'Pendaftaran mandiri; menunggu approval admin',
        },
      },
    });

    return NextResponse.json(
      {
        success: true,
        message: 'Pendaftaran berhasil. Akun Anda sedang menunggu persetujuan dari Administrator.',
        userId: newUser.id,
      },
      { status: 201 }
    );
  } catch (err: any) {
    console.error('[API_REGISTER_ERROR]', err);
    return NextResponse.json(
      {
        success: false,
        error: 'Terjadi kesalahan internal saat memproses pendaftaran.',
      },
      { status: 500 }
    );
  }
}
