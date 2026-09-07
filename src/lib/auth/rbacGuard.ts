import { NextResponse } from 'next/server';
import { extractSessionToken, validateSessionToken, AuthenticatedUser } from './session';

// ==============================================================================
// SERVER-SIDE RBAC GUARD & AUTHORIZATION HELPERS
// Ensures all access decisions happen securely on the backend
// ROOT_ADMIN holds system-wide override; Granular permissions enforced for all else
// ==============================================================================

export class AuthError extends Error {
  public statusCode: number;
  constructor(message: string, statusCode = 401) {
    super(message);
    this.name = 'AuthError';
    this.statusCode = statusCode;
  }
}

// ==============================================================================
// CSRF ORIGIN VALIDATION (Configured Trusted Origin)
// ==============================================================================

/**
 * Validates request Origin against configured trusted origins (APP_ORIGIN / NEXT_PUBLIC_APP_URL).
 * Naive Host-header trust is explicitly prohibited.
 * Pure Bearer-token requests (without cookies) are exempt from CSRF.
 */
export function verifyCsrfOrigin(req: Request): boolean {
  const method = req.method.toUpperCase();
  // Safe HTTP methods do not require CSRF validation
  if (['GET', 'HEAD', 'OPTIONS'].includes(method)) {
    return true;
  }

  const authHeader = req.headers.get('authorization');
  const cookieHeader = req.headers.get('cookie');
  const hasSessionCookie = cookieHeader && cookieHeader.includes('__session=');

  // Pure Bearer-token requests without browser cookies are exempt from CSRF
  if (authHeader && authHeader.startsWith('Bearer ') && !hasSessionCookie) {
    return true;
  }

  // Cross-site fetch metadata check
  const secFetchSite = req.headers.get('sec-fetch-site');
  if (secFetchSite === 'cross-site') {
    return false;
  }

  // Extract Origin from header, fallback to Referer origin
  const originHeader = req.headers.get('origin');
  let requestOrigin = originHeader;
  if (!requestOrigin) {
    const referer = req.headers.get('referer');
    if (referer) {
      try {
        requestOrigin = new URL(referer).origin;
      } catch {
        requestOrigin = null;
      }
    }
  }

  // If no Origin or Referer header was provided:
  if (!requestOrigin) {
    // In production, state-changing requests with cookie MUST have an Origin or Referer
    if (process.env.NODE_ENV === 'production') {
      return false;
    }
    // In development / test, allow CLI/scripts if no cross-site flag is present
    return true;
  }

  // Configured trusted origins list
  const configuredOrigins = [
    process.env.APP_ORIGIN,
    process.env.NEXT_PUBLIC_APP_URL,
    'http://localhost:3000',
    'http://127.0.0.1:3000',
  ].filter(Boolean) as string[];

  const normalizedRequestOrigin = requestOrigin.replace(/\/+$/, '').toLowerCase();
  return configuredOrigins.some((allowed) => allowed.replace(/\/+$/, '').toLowerCase() === normalizedRequestOrigin);
}

/**
 * Enforces CSRF origin check or throws AuthError with 403 status.
 */
export function checkCsrf(req: Request): void {
  if (!verifyCsrfOrigin(req)) {
    throw new AuthError('Akses ditolak: Verifikasi CSRF origin gagal.', 403);
  }
}

/**
 * Authenticates request without throwing. Returns AuthenticatedUser or null.
 */
export async function authenticateRequest(req: Request): Promise<AuthenticatedUser | null> {
  const token = extractSessionToken(req);
  if (!token) {
    return null;
  }

  return await validateSessionToken(token);
}

/**
 * Requires valid active session. Throws 401 if missing or invalid.
 */
export async function requireAuth(req: Request): Promise<AuthenticatedUser> {
  const user = await authenticateRequest(req);
  if (!user) {
    throw new AuthError('Otentikasi diperlukan untuk mengakses resource ini.', 401);
  }
  return user;
}

/**
 * Requires specific granular permission code.
 * ROOT_ADMIN automatically bypasses all granular restrictions.
 */
export async function requirePermission(req: Request, permissionCode: string): Promise<AuthenticatedUser> {
  const user = await requireAuth(req);

  // ROOT_ADMIN has full system access
  if (user.isSystemRoot || user.roles.includes('ROOT_ADMIN')) {
    return user;
  }

  // Check granular permission
  if (!user.permissions.includes(permissionCode)) {
    throw new AuthError(`Akses ditolak: Memerlukan izin '${permissionCode}'.`, 403);
  }

  return user;
}

/**
 * Requires System Root Administrator.
 */
export async function requireRootAdmin(req: Request): Promise<AuthenticatedUser> {
  const user = await requireAuth(req);

  if (!user.isSystemRoot && !user.roles.includes('ROOT_ADMIN')) {
    throw new AuthError('Akses ditolak: Hanya Root Administrator yang berwenang.', 403);
  }

  return user;
}

/**
 * Standardized JSON response for unauthenticated access.
 */
export function unauthorizedResponse(message = 'Otentikasi diperlukan.'): NextResponse {
  return NextResponse.json(
    { success: false, error: message, code: 'UNAUTHORIZED' },
    { status: 401 }
  );
}

/**
 * Standardized JSON response for forbidden access.
 */
export function forbiddenResponse(message = 'Akses ditolak.'): NextResponse {
  return NextResponse.json(
    { success: false, error: message, code: 'FORBIDDEN' },
    { status: 403 }
  );
}

/**
 * Standardized JSON response for bad request / validation error.
 */
export function badRequestResponse(message: string, details?: any): NextResponse {
  return NextResponse.json(
    { success: false, error: message, code: 'BAD_REQUEST', details },
    { status: 400 }
  );
}
