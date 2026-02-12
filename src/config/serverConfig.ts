// 최소 설정: API(5500)와 RAG(5600)만 사용
export interface ServerConfig {
  api: { url: string };
  rag: { url: string };
}

export function getServerConfig(): ServerConfig {
  return {
    api: {
      url: process.env.NEXT_PUBLIC_API_SERVER_URL || 'http://localhost:5500/api/v1',
    },
    rag: {
      url: process.env.NEXT_PUBLIC_RAG_API_SERVER || 'http://localhost:5600',
    },
  };
}

export const serverConfig = getServerConfig();

export const getApiServerUrl = (): string => {
  const isServer = typeof window === 'undefined';
  if (isServer && process.env.API_SERVER_URL) return process.env.API_SERVER_URL;
  return process.env.NEXT_PUBLIC_API_SERVER_URL || 'http://localhost:5500/api/v1';
};

export const getRagServerUrl = (): string => {
  const isServer = typeof window === 'undefined';
  if (isServer && process.env.RAG_API_SERVER) return process.env.RAG_API_SERVER;
  return process.env.NEXT_PUBLIC_RAG_API_SERVER || 'http://localhost:5600';
};
