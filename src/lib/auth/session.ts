import crypto from 'crypto';
import { prisma } from '@/lib/db/prisma';
import { UserStatus, Prisma } from '@prisma/client';

// ==============================================================================
// ENTERPRISE SESSION MANAGEMENT
// - Opaque 256-bit cryptographically random token
// - Token hashed with SHA-256 before storage in PostgreSQL
// - Redis cache layer (15-min sliding TTL) with fail-closed PostgreSQL fallback
// - HttpOnly, Secure, SameSite=Lax cookie delivery
// ==============================================================================

export const SESSION_COOKIE_NAME = '__session';
export const SESSION_TTL_SECONDS = 7 * 24 * 60 * 60; // 7 Days
export const REDIS_SESSION_TTL_SECONDS = 15 * 60; // 15 Minutes cache

export interface AuthenticatedUser {
  id: string;
  username: string;
  email: string;
  fullName: string;
  status: UserStatus;
  isSystemRoot: boolean;
  roles: string[];
  permissions: string[];
}

export interface AuthoritativeSessionRow {
  session_id: string;
  user_id: string;
  expires_at: Date;
  caller_id: string;
  status: UserStatus;
  is_system_root: boolean;
}

/**
 * Computes deterministic SHA-256 hash of an opaque session token.
 */
export function hashSessionToken(token: string): string {
  return crypto.createHash('sha256').update(token).digest('hex');
}

/**
 * Creates a new session in PostgreSQL.
 * Redis is NOT used for authorization or session caching.
 * Supports tx for Transaction Client Confinement.
 */
export async function createSession(
  userId: string,
  ipAddress?: string,
  userAgent?: string,
  tx?: Prisma.TransactionClient
): Promise<{ sessionToken: string; expiresAt: Date }> {
  // Generate 256-bit secure random token
  const sessionToken = crypto.randomBytes(32).toString('hex');
  const tokenHash = hashSessionToken(sessionToken);
  const expiresAt = new Date(Date.now() + SESSION_TTL_SECONDS * 1000);

  // Authoritative write to PostgreSQL
  const client = tx || prisma;
  await client.session.create({
    data: {
      userId,
      tokenHash,
      ipAddress: ipAddress ? ipAddress.slice(0, 45) : null,
      userAgent: userAgent ? userAgent.slice(0, 500) : null,
      expiresAt,
    },
  });

  return { sessionToken, expiresAt };
}

/**
 * Authoritative session lookup with parent user row lock (FOR UPDATE OF u).
 * Locks 'users' exclusively while leaving 'sessions' lock-free.
 * Enforces Model A session expiry: s.expires_at > clock_timestamp().
 */
export async function getAuthoritativeSessionWithUserLock(
  tx: Prisma.TransactionClient,
  sessionTokenHash: string
): Promise<AuthoritativeSessionRow | null> {
  const rows: any[] = await tx.$queryRawUnsafe(
    `SELECT
        s.id AS session_id,
        s.user_id,
        s.expires_at,
        u.id AS caller_id,
        u.status,
        u.is_system_root
     FROM sessions s
     JOIN users u ON u.id = s.user_id
     WHERE s.token_hash = $1
       AND s.expires_at > clock_timestamp()
     FOR UPDATE OF u;`,
    sessionTokenHash
  );

  if (!rows || rows.length === 0) return null;
  return rows[0];
}

/**
 * Validates session token authoritatively against PostgreSQL.
 * Single unified indexed SQL query resolves session existence, account status,
 * isSystemRoot, roles, and permissions under one coherent snapshot.
 * Redis is completely bypassed for authentication and authorization.
 */
export async function validateSessionToken(token: string): Promise<AuthenticatedUser | null> {
  if (!token || typeof token !== 'string' || token.length < 32) {
    return null;
  }

  const tokenHash = hashSessionToken(token);

  // Authoritative, indexed, single-round-trip query against PostgreSQL
  const rows: any[] = await prisma.$queryRawUnsafe(
    `SELECT s.id AS session_id, s.expires_at,
            u.id AS user_id, u.username, u.email, u.full_name, u.status, u.is_system_root,
            COALESCE(array_agg(DISTINCT r.name) FILTER (WHERE r.name IS NOT NULL), '{}') AS roles,
            COALESCE(array_agg(DISTINCT p.code) FILTER (WHERE p.code IS NOT NULL), '{}') AS permissions
     FROM sessions s
     JOIN users u ON u.id = s.user_id
     LEFT JOIN user_roles ur ON ur.user_id = u.id
     LEFT JOIN roles r ON r.id = ur.role_id
     LEFT JOIN role_permissions rp ON rp.role_id = r.id
     LEFT JOIN permissions p ON p.id = rp.permission_id
     WHERE s.token_hash = $1
       AND s.expires_at > clock_timestamp()
       AND u.status = 'ACTIVE'
     GROUP BY s.id, u.id;`,
    tokenHash
  );

  if (!rows || rows.length === 0) {
    return null;
  }

  const row = rows[0];

  return {
    id: row.user_id,
    username: row.username,
    email: row.email,
    fullName: row.full_name,
    status: row.status as UserStatus,
    isSystemRoot: Boolean(row.is_system_root),
    roles: Array.isArray(row.roles) ? row.roles : [],
    permissions: Array.isArray(row.permissions) ? row.permissions : [],
  };
}

/**
 * Revokes a session authoritatively from PostgreSQL.
 */
export async function revokeSession(token: string): Promise<void> {
  if (!token) return;
  const tokenHash = hashSessionToken(token);

  try {
    await prisma.session.deleteMany({
      where: { tokenHash },
    });
  } catch (dbErr) {
    console.warn('[SESSION_REVOKE_WARNING] PostgreSQL delete error:', dbErr);
  }
}

/**
 * Deterministically invalidates ALL active sessions for a user across PostgreSQL.
 * Enforces Universal Lock Hierarchy: users -> sessions.
 */
export async function invalidateUserSessions(userId: string): Promise<void> {
  if (!userId) return;

  try {
    await prisma.$transaction(async (tx) => {
      await tx.$queryRawUnsafe(`SELECT id FROM users WHERE id = $1::uuid FOR UPDATE`, userId);
      await tx.session.deleteMany({
        where: { userId },
      });
    });
  } catch (dbErr) {
    console.warn('[USER_SESSIONS_REVOKE_WARNING] PostgreSQL deleteMany error:', dbErr);
  }
}

/**
 * Updates user status in PostgreSQL with Model B parent user row locking.
 * If status transitions to non-ACTIVE (DISABLED, PENDING, SUSPENDED),
 * all active sessions in DB are immediately purged in the same transaction.
 */
export async function setUserStatus(userId: string, status: UserStatus): Promise<void> {
  await prisma.$transaction(async (tx) => {
    // Lock user row first under Model B
    await tx.$queryRawUnsafe(`SELECT id FROM users WHERE id = $1::uuid FOR UPDATE`, userId);

    await tx.user.update({
      where: { id: userId },
      data: { status },
    });

    if (status !== UserStatus.ACTIVE) {
      await tx.session.deleteMany({
        where: { userId },
      });
    }
  });
}

/**
 * Extracts session token from incoming Request (Cookies or Bearer Authorization).
 */
export function extractSessionToken(req: Request): string | null {
  // 1. Check Cookie header
  const cookieHeader = req.headers.get('cookie');
  if (cookieHeader) {
    const cookies = cookieHeader.split(';').map((c) => c.trim());
    for (const c of cookies) {
      if (c.startsWith(`${SESSION_COOKIE_NAME}=`)) {
        const val = c.substring(SESSION_COOKIE_NAME.length + 1);
        if (val) return decodeURIComponent(val);
      }
    }
  }

  // 2. Check Authorization header: Bearer <token>
  const authHeader = req.headers.get('authorization');
  if (authHeader && authHeader.startsWith('Bearer ')) {
    const val = authHeader.substring(7).trim();
    if (val) return val;
  }

  return null;
}

/**
 * Builds Set-Cookie string for session establishment.
 */
export function buildSessionCookie(token: string, expiresAt: Date): string {
  const isProd = process.env.NODE_ENV === 'production';
  const parts = [
    `${SESSION_COOKIE_NAME}=${encodeURIComponent(token)}`,
    `Path=/`,
    `Expires=${expiresAt.toUTCString()}`,
    `Max-Age=${SESSION_TTL_SECONDS}`,
    `HttpOnly`,
    `SameSite=Lax`,
  ];

  if (isProd) {
    parts.push('Secure');
  }

  return parts.join('; ');
}

/**
 * Builds Set-Cookie string for clearing the session on logout.
 */
export function buildLogoutCookie(): string {
  const isProd = process.env.NODE_ENV === 'production';
  const parts = [
    `${SESSION_COOKIE_NAME}=`,
    `Path=/`,
    `Expires=Thu, 01 Jan 1970 00:00:00 GMT`,
    `Max-Age=0`,
    `HttpOnly`,
    `SameSite=Lax`,
  ];

  if (isProd) {
    parts.push('Secure');
  }

  return parts.join('; ');
}
