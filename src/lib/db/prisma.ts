import { PrismaClient } from '@prisma/client';

// ==============================================================================
// PRISMA CLIENT SINGLETON (BOUNDED CONNECTION POOL)
// Prevents connection exhaustion in Next.js development hot-reloading
// Connection pool size: C_web = 12 (configured via DATABASE_URL connection_limit)
// ==============================================================================

const prismaClientSingleton = () => {
  return new PrismaClient({
    log:
      process.env.NODE_ENV === 'development'
        ? ['warn', 'error']
        : ['error'],
  });
};

type PrismaClientSingleton = ReturnType<typeof prismaClientSingleton>;

const globalForPrisma = globalThis as unknown as {
  prisma: PrismaClientSingleton | undefined;
};

export const prisma = globalForPrisma.prisma ?? prismaClientSingleton();

if (process.env.NODE_ENV !== 'production') {
  globalForPrisma.prisma = prisma;
}

export default prisma;
