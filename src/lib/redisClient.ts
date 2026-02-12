import Redis from 'ioredis';

let client: Redis | null = null;

export function getRedisClient() {
  if (client) return client;

  const url = process.env.REDIS_URL || process.env.NEXT_PUBLIC_REDIS_URL;
  if (!url) return null;

  client = new Redis(url, {
    maxRetriesPerRequest: 1,
    enableReadyCheck: false,
    lazyConnect: true
  });

  // 연결 시도 지연: 처음 요청 시 connect 호출
  return client;
}
