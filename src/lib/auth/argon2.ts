import { hash, verify } from '@node-rs/argon2';

// ==============================================================================
// ARGON2ID PASSWORD HASHING HELPER
// Parameters adhere to OWASP / Enterprise Security Guidelines:
// - Default Algorithm: Argon2id ($argon2id$v=19)
// - Memory Cost: 65536 KB (64 MB)
// - Iterations (timeCost): 3
// - Parallelism: 4
// - Output Length: 32 bytes
// ==============================================================================

export async function hashPassword(password: string): Promise<string> {
  if (!password || typeof password !== 'string') {
    throw new Error('Password must be a non-empty string');
  }

  return hash(password, {
    memoryCost: 65536,
    timeCost: 3,
    outputLen: 32,
    parallelism: 4,
  });
}

export async function verifyPassword(passwordHash: string, plainText: string): Promise<boolean> {
  if (!passwordHash || !plainText) {
    return false;
  }

  try {
    return await verify(passwordHash, plainText);
  } catch (err) {
    console.error('[ARGON2] Verification error:', err);
    return false;
  }
}
