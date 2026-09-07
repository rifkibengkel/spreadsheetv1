import { redis, isRedisHealthy } from '@/lib/redis/redisClient';

// ==============================================================================
// MULTI-DIMENSIONAL AUTHENTICATION RATE LIMITER
// - Prevents User A failed attempts from blocking User B (Per-Account Isolation)
// - Defends against distributed brute-force attacks (Per-IP & Per-Account dimensions)
// - Window: 10 minutes (600s), Max 5 failed attempts per account, Max 20 per IP
// - Backed by Redis sliding window counter; fails gracefully if Redis unavailable
// ==============================================================================

const ACCOUNT_MAX_ATTEMPTS = 5;
const IP_MAX_ATTEMPTS = 20;
const WINDOW_SECONDS = 600; // 10 minutes

export interface RateLimitResult {
  allowed: boolean;
  retryAfterSeconds?: number;
  reason?: string;
}

/**
 * Normalizes username or email to prevent casing bypasses.
 */
function normalizeIdentifier(id: string): string {
  return id.trim().toLowerCase();
}

/**
 * Checks both Account-level and IP-level rate limits.
 */
export async function checkAuthRateLimit(
  accountIdentifier: string,
  clientIp: string
): Promise<RateLimitResult> {
  try {
    if (!(await isRedisHealthy())) {
      // Graceful fallback: Allow request, relied upon Argon2id high memory/CPU cost
      return { allowed: true };
    }

    const normAccount = normalizeIdentifier(accountIdentifier);
    const accountKey = `ratelimit:auth:account:${normAccount}`;
    const ipKey = `ratelimit:auth:ip:${clientIp}`;

    const [accountAttemptsStr, ipAttemptsStr, accountTtl, ipTtl] = await Promise.all([
      redis.get(accountKey),
      redis.get(ipKey),
      redis.ttl(accountKey),
      redis.ttl(ipKey),
    ]);

    const accountAttempts = accountAttemptsStr ? parseInt(accountAttemptsStr, 10) : 0;
    const ipAttempts = ipAttemptsStr ? parseInt(ipAttemptsStr, 10) : 0;

    // 1. Check Account threshold
    if (accountAttempts >= ACCOUNT_MAX_ATTEMPTS) {
      const waitTime = accountTtl > 0 ? accountTtl : WINDOW_SECONDS;
      return {
        allowed: false,
        retryAfterSeconds: waitTime,
        reason: `Terlalu banyak percobaan gagal pada akun ini. Silakan coba lagi dalam ${Math.ceil(waitTime / 60)} menit.`,
      };
    }

    // 2. Check IP threshold
    if (ipAttempts >= IP_MAX_ATTEMPTS) {
      const waitTime = ipTtl > 0 ? ipTtl : WINDOW_SECONDS;
      return {
        allowed: false,
        retryAfterSeconds: waitTime,
        reason: `Terlalu banyak percobaan login dari jaringan ini. Silakan coba lagi dalam ${Math.ceil(waitTime / 60)} menit.`,
      };
    }

    return { allowed: true };
  } catch (err) {
    console.warn('[RATELIMIT_WARNING] Rate limit check error:', err);
    return { allowed: true }; // Fail-open on unexpected Redis error
  }
}

/**
 * Records an authentication failure for both Account and IP.
 */
export async function recordAuthFailure(
  accountIdentifier: string,
  clientIp: string
): Promise<void> {
  try {
    if (!(await isRedisHealthy())) return;

    const normAccount = normalizeIdentifier(accountIdentifier);
    const accountKey = `ratelimit:auth:account:${normAccount}`;
    const ipKey = `ratelimit:auth:ip:${clientIp}`;

    const multi = redis.multi();

    // Increment account counter
    multi.incr(accountKey);
    multi.expire(accountKey, WINDOW_SECONDS);

    // Increment IP counter
    multi.incr(ipKey);
    multi.expire(ipKey, WINDOW_SECONDS);

    await multi.exec();
  } catch (err) {
    console.warn('[RATELIMIT_WARNING] Failed to record auth failure:', err);
  }
}

/**
 * Clears failed attempts upon successful authentication.
 */
export async function clearAuthFailure(
  accountIdentifier: string,
  clientIp: string
): Promise<void> {
  try {
    if (!(await isRedisHealthy())) return;

    const normAccount = normalizeIdentifier(accountIdentifier);
    const accountKey = `ratelimit:auth:account:${normAccount}`;
    const ipKey = `ratelimit:auth:ip:${clientIp}`;

    await redis.del(accountKey, ipKey);
  } catch {
    // Non-fatal
  }
}

/**
 * TRUSTED REVERSE PROXY & CLIENT IP EXTRACTION
 *
 * Trust Model & Security Invariants:
 * 1. Reverse Proxy Boundary:
 *    - In production, Next.js runs behind a trusted reverse proxy / edge
 *      (e.g., Cloudflare, AWS ALB, Nginx).
 *    - When TRUST_PROXY=true (or '1'):
 *        a. Prioritizes 'cf-connecting-ip' (Cloudflare Edge, immune to client spoofing).
 *        b. Prioritizes 'x-real-ip' (Nginx/ALB configured to set $remote_addr).
 *        c. For 'x-forwarded-for': Parses comma-separated hops, trims whitespace.
 *    - When TRUST_PROXY is false/undefined and not in production:
 *        a. Client-supplied headers (X-Forwarded-For, X-Real-IP) CANNOT BE TRUSTED,
 *           as an attacker on a public interface could spoof headers to bypass IP rate limits.
 *        b. Falls back safely to '127.0.0.1'.
 *
 * 2. Key Injection Sanitization:
 *    - Strips IPv6-mapped IPv4 prefix ("::ffff:").
 *    - Verifies safe characters [0-9a-fA-F:.] to prevent Redis delimiter/key injection.
 */
export function getClientIp(req: Request): string {
  const trustProxy =
    process.env.TRUST_PROXY === 'true' ||
    process.env.TRUST_PROXY === '1' ||
    process.env.NODE_ENV === 'production';

  let rawIp = '';

  if (trustProxy) {
    // 1. Cloudflare edge header (cannot be spoofed if traffic routes through Cloudflare)
    const cfIp = req.headers.get('cf-connecting-ip');
    if (cfIp) {
      rawIp = cfIp.trim();
    } else {
      // 2. Nginx/ALB trusted single remote address
      const realIp = req.headers.get('x-real-ip');
      if (realIp) {
        rawIp = realIp.trim();
      } else {
        // 3. X-Forwarded-For chain: client, proxy1, proxy2...
        const forwarded = req.headers.get('x-forwarded-for');
        if (forwarded) {
          const parts = forwarded.split(',');
          // First non-empty entry represents originating client as reported by proxy
          rawIp = parts[0].trim();
        }
      }
    }
  }

  if (!rawIp) {
    rawIp = '127.0.0.1';
  }

  // Sanitize IPv6-mapped IPv4 addresses (e.g., ::ffff:192.168.1.1 -> 192.168.1.1)
  if (rawIp.startsWith('::ffff:')) {
    rawIp = rawIp.substring(7);
  }

  // Defensive sanitization: verify IP contains only valid IPv4/IPv6 characters to prevent Redis key injection
  if (!/^[0-9a-fA-F:.]+$/.test(rawIp) || rawIp.length > 45) {
    return '127.0.0.1';
  }

  return rawIp;
}
