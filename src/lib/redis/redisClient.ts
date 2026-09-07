import Redis, { RedisOptions } from 'ioredis';

// ==============================================================================
// REDIS CLIENT SINGLETON (LOCAL SIMULATION OF PRODUCTION TOPOLOGY)
// Manages BullMQ Queues, Sliding Window Rate Limiting, & Pub/Sub Progress
// Uses globalThis singleton pattern to prevent connection leaks during Next.js HMR
// ==============================================================================

const REDIS_URL = process.env.REDIS_URL || 'redis://127.0.0.1:6379';

const redisOptions: RedisOptions = {
  maxRetriesPerRequest: null, // Required by BullMQ
  enableReadyCheck: true,
  lazyConnect: true,
  retryStrategy(times) {
    // Exponential backoff with a maximum delay of 3 seconds
    const delay = Math.min(times * 100, 3000);
    return delay;
  },
  reconnectOnError(err) {
    const targetError = 'READONLY';
    if (err.message.includes(targetError)) {
      return true; // Reconnect when replica becomes master
    }
    return false;
  },
};

function createRedisClient(): Redis {
  const client = new Redis(REDIS_URL, redisOptions);

  client.on('error', (err) => {
    // Log error cleanly without crashing Next.js or worker process
    console.warn('[REDIS_CLIENT_WARNING] Redis connection issue:', err.message);
  });

  client.on('connect', () => {
    if (process.env.NODE_ENV !== 'test') {
      console.log('[REDIS] Successfully connected to Redis server');
    }
  });

  return client;
}

const globalForRedis = globalThis as unknown as {
  redisClient: Redis | undefined;
};

export const redis = globalForRedis.redisClient ?? createRedisClient();

if (process.env.NODE_ENV !== 'production') {
  globalForRedis.redisClient = redis;
}

/**
 * Health check helper for Redis availability.
 * Returns true if Redis is reachable and answers PING.
 */
export async function isRedisHealthy(): Promise<boolean> {
  try {
    if (redis.status === 'wait') {
      await redis.connect();
    }
    const pong = await redis.ping();
    return pong === 'PONG';
  } catch {
    return false;
  }
}

export default redis;
