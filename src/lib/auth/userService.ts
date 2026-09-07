import { prisma } from '@/lib/db/prisma';
import { UserStatus } from '@prisma/client';
import { AuthError } from '@/lib/auth/rbacGuard';

export interface UpdateUserPayload {
  fullName?: string;
  status?: UserStatus;
  isSystemRoot?: boolean;
  roleNames?: string[];
}

export interface UpdateUserParams {
  sessionTokenHash: string;
  targetUserId: string;
  payload: UpdateUserPayload;
}

export interface DeleteUserParams {
  sessionTokenHash: string;
  targetUserId: string;
}

export class AuthForbiddenError extends Error {
  statusCode: number = 403;
  constructor(message: string) {
    super(message);
    this.name = 'AuthForbiddenError';
  }
}

export class AuthNotFoundError extends Error {
  statusCode: number = 404;
  constructor(message: string) {
    super(message);
    this.name = 'AuthNotFoundError';
  }
}

const MAX_RETRIES = 3;

/**
 * Updates a user account with strict Model B Universal Lock Protocol and Root Admin Invariant enforcement.
 * 
 * TRANSACTIONAL INVARIANTS:
 * 1. Transaction Client Confinement: All operations strictly execute within tx.
 * 2. Universal Lock Order: Caller and target user rows are locked in canonical ascending UUID order.
 * 3. Authoritative Session Resolution: Caller identity and permissions are resolved from PostgreSQL inside tx.
 * 4. Model B Child Mutation Protection: Target user row is locked before any user_roles mutation.
 * 5. Bounded Retry: Automatically retries on SQLSTATE 40001 (serialization) and 40P01 (deadlock).
 */
export async function updateUser(params: UpdateUserParams): Promise<void> {
  const { sessionTokenHash, targetUserId, payload } = params;

  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      await prisma.$transaction(async (tx) => {
        // 1. Authoritative session lookup & caller identity resolution
        const sessionRows: any[] = await tx.$queryRawUnsafe(
          `SELECT s.id AS session_id, s.user_id, s.expires_at 
           FROM sessions s 
           WHERE s.token_hash = $1 AND s.expires_at > clock_timestamp()`,
          sessionTokenHash
        );

        if (!sessionRows || sessionRows.length === 0) {
          throw new AuthError('Sesi tidak ditemukan atau telah kedaluwarsa.', 401);
        }

        const callerUserId: string = sessionRows[0].user_id;

        // 2. Canonical User Row Lock Hierarchy (Deadlock-Free: min(UUID) -> max(UUID))
        const distinctUserIds = Array.from(new Set([callerUserId, targetUserId])).sort();
        for (const uid of distinctUserIds) {
          await tx.$queryRawUnsafe(`SELECT id FROM users WHERE id = $1::uuid FOR UPDATE`, uid);
        }

        // 3. Re-verify session validity in tx after lock acquisition
        const validSession: any[] = await tx.$queryRawUnsafe(
          `SELECT id FROM sessions WHERE token_hash = $1 AND expires_at > clock_timestamp()`,
          sessionTokenHash
        );
        if (!validSession || validSession.length === 0) {
          throw new AuthError('Sesi telah dicabut sebelum transaksi dimulai.', 401);
        }

        // 4. Fetch caller user record under lock
        const callerRecord = await tx.user.findUnique({
          where: { id: callerUserId },
          include: {
            roles: {
              include: {
                role: {
                  include: {
                    permissions: {
                      include: { permission: true },
                    },
                  },
                },
              },
            },
          },
        });

        if (!callerRecord || callerRecord.status !== UserStatus.ACTIVE) {
          throw new AuthError('Akun caller tidak aktif atau telah dinonaktifkan.', 401);
        }

        // Compile caller permissions
        const callerPermissions = new Set<string>();
        for (const ur of callerRecord.roles) {
          for (const rp of ur.role.permissions) {
            callerPermissions.add(rp.permission.code);
          }
        }

        // Check permission: USER_MANAGE or isSystemRoot / ROOT_ADMIN
        const isCallerRoot = callerRecord.isSystemRoot || callerRecord.roles.some((r) => r.role.name === 'ROOT_ADMIN');
        if (!isCallerRoot && !callerPermissions.has('USER_MANAGE')) {
          throw new AuthForbiddenError('Akses ditolak: Memerlukan izin USER_MANAGE.');
        }

        // 5. Fetch target user record under lock
        const targetUser = await tx.user.findUnique({
          where: { id: targetUserId },
          include: {
            roles: {
              include: { role: true },
            },
          },
        });

        if (!targetUser) {
          throw new AuthNotFoundError('User tidak ditemukan.');
        }

        // 6. Root Administrator Invariant Enforcement
        if (targetUser.isSystemRoot) {
          if (!callerRecord.isSystemRoot) {
            throw new AuthForbiddenError('Akun root administrator tidak dapat dimodifikasi oleh admin biasa.');
          }
          if (payload.isSystemRoot === false) {
            throw new AuthForbiddenError('Flag system-root tidak dapat dihapus dari akun root administrator.');
          }
          if (payload.roleNames && !payload.roleNames.includes('ROOT_ADMIN')) {
            throw new AuthForbiddenError('Role ROOT_ADMIN pada akun root administrator tidak dapat dicabut (demoted).');
          }
        }

        // Non-root caller cannot grant root privileges
        if (!callerRecord.isSystemRoot) {
          if (payload.isSystemRoot === true) {
            throw new AuthForbiddenError('Hanya system root yang dapat mengangkat user menjadi system root.');
          }
          if (payload.roleNames && payload.roleNames.includes('ROOT_ADMIN')) {
            throw new AuthForbiddenError('Hanya system root yang dapat memberikan role ROOT_ADMIN.');
          }
        }

        // 7. Self-Modification Protections
        if (callerUserId === targetUserId) {
          if (payload.status && payload.status !== UserStatus.ACTIVE) {
            throw new AuthForbiddenError('Tidak dapat menonaktifkan akun sendiri.');
          }
          if (callerRecord.isSystemRoot && payload.roleNames && !payload.roleNames.includes('ROOT_ADMIN')) {
            throw new AuthForbiddenError('Tidak dapat mencabut role root administrator sendiri.');
          }
        }

        // 8. Apply User updates using tx
        const updateData: any = {};
        if (payload.fullName !== undefined) updateData.fullName = payload.fullName;
        if (payload.status !== undefined) updateData.status = payload.status;
        if (payload.isSystemRoot !== undefined && callerRecord.isSystemRoot) {
          updateData.isSystemRoot = payload.isSystemRoot;
        }

        if (Object.keys(updateData).length > 0) {
          await tx.user.update({
            where: { id: targetUserId },
            data: updateData,
          });
        }

        // 9. Apply Role updates using tx (under target user lock)
        if (payload.roleNames) {
          const roles = await tx.role.findMany({
            where: { name: { in: payload.roleNames } },
          });

          // Delete existing roles
          await tx.userRole.deleteMany({
            where: { userId: targetUserId },
          });

          // Insert new roles
          for (const r of roles) {
            await tx.userRole.create({
              data: {
                userId: targetUserId,
                roleId: r.id,
              },
            });
          }
        }

        // 10. Purge active sessions if status transitions to non-ACTIVE
        if (payload.status && payload.status !== UserStatus.ACTIVE) {
          await tx.session.deleteMany({
            where: { userId: targetUserId },
          });
        }
      });

      return;
    } catch (err: any) {
      if ((err.code === '40001' || err.code === '40P01' || err.message?.includes('deadlock') || err.message?.includes('could not serialize')) && attempt < MAX_RETRIES) {
        const delay = Math.pow(2, attempt) * 20 + Math.random() * 20;
        await new Promise((r) => setTimeout(r, delay));
        continue;
      }
      throw err;
    }
  }
}

/**
 * Deletes a user account with strict Model B Universal Lock Protocol and Root Admin Invariant enforcement.
 */
export async function deleteUser(params: DeleteUserParams): Promise<void> {
  const { sessionTokenHash, targetUserId } = params;

  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      await prisma.$transaction(async (tx) => {
        // 1. Authoritative session lookup
        const sessionRows: any[] = await tx.$queryRawUnsafe(
          `SELECT s.id AS session_id, s.user_id 
           FROM sessions s 
           WHERE s.token_hash = $1 AND s.expires_at > clock_timestamp()`,
          sessionTokenHash
        );

        if (!sessionRows || sessionRows.length === 0) {
          throw new AuthError('Sesi tidak ditemukan atau telah kedaluwarsa.', 401);
        }

        const callerUserId: string = sessionRows[0].user_id;

        if (callerUserId === targetUserId) {
          throw new AuthForbiddenError('Tidak dapat menghapus akun sendiri.');
        }

        // 2. Canonical User Row Lock Hierarchy (Deadlock-Free: min(UUID) -> max(UUID))
        const distinctUserIds = Array.from(new Set([callerUserId, targetUserId])).sort();
        for (const uid of distinctUserIds) {
          await tx.$queryRawUnsafe(`SELECT id FROM users WHERE id = $1::uuid FOR UPDATE`, uid);
        }

        // 3. Re-verify session validity in tx
        const validSession: any[] = await tx.$queryRawUnsafe(
          `SELECT id FROM sessions WHERE token_hash = $1 AND expires_at > clock_timestamp()`,
          sessionTokenHash
        );
        if (!validSession || validSession.length === 0) {
          throw new AuthError('Sesi telah dicabut sebelum transaksi dimulai.', 401);
        }

        // 4. Fetch caller user record under lock
        const callerRecord = await tx.user.findUnique({
          where: { id: callerUserId },
          include: {
            roles: {
              include: {
                role: {
                  include: {
                    permissions: {
                      include: { permission: true },
                    },
                  },
                },
              },
            },
          },
        });

        if (!callerRecord || callerRecord.status !== UserStatus.ACTIVE) {
          throw new AuthError('Akun caller tidak aktif atau telah dinonaktifkan.', 401);
        }

        const callerPermissions = new Set<string>();
        for (const ur of callerRecord.roles) {
          for (const rp of ur.role.permissions) {
            callerPermissions.add(rp.permission.code);
          }
        }

        const isCallerRoot = callerRecord.isSystemRoot || callerRecord.roles.some((r) => r.role.name === 'ROOT_ADMIN');
        if (!isCallerRoot && !callerPermissions.has('USER_MANAGE')) {
          throw new AuthForbiddenError('Akses ditolak: Memerlukan izin USER_MANAGE.');
        }

        // 5. Fetch target user record under lock
        const targetUser = await tx.user.findUnique({
          where: { id: targetUserId },
        });

        if (!targetUser) {
          throw new AuthNotFoundError('User tidak ditemukan.');
        }

        // Root Admin CANNOT be deleted
        if (targetUser.isSystemRoot) {
          throw new AuthForbiddenError('Akun root administrator bersifat permanen dan tidak dapat dihapus.');
        }

        // 6. Delete all sessions for target user first
        await tx.session.deleteMany({
          where: { userId: targetUserId },
        });

        // 7. Delete target user from PostgreSQL (Cascades to user_roles, audit_logs)
        await tx.user.delete({
          where: { id: targetUserId },
        });
      });

      return;
    } catch (err: any) {
      if ((err.code === '40001' || err.code === '40P01' || err.message?.includes('deadlock') || err.message?.includes('could not serialize')) && attempt < MAX_RETRIES) {
        const delay = Math.pow(2, attempt) * 20 + Math.random() * 20;
        await new Promise((r) => setTimeout(r, delay));
        continue;
      }
      throw err;
    }
  }
}
