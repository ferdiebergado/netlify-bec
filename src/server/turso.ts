// eslint-disable-next-line import/no-extraneous-dependencies
import { createClient } from '@libsql/client';
import config from './config';

export default function useTurso() {
  const { uri, token } = config.db;

  if (!uri) throw new Error('VITE_TURSO_DB_URL not set!');

  return createClient({
    url: uri,
    authToken: token,
  });
}
