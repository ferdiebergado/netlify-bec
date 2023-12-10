// eslint-disable-next-line import/no-extraneous-dependencies
import { createClient, Client } from '@libsql/client';
import config from './config';

/**
 * Creates a Turso database client based on the configuration settings.
 *
 * @throws {Error} Throws an error if the database URL or authentication token is not set.
 * @returns {Client} The Turso database client.
 */
export default function useTurso(): Client {
  const { uri, token } = config.db;

  if (!uri || !token) throw new Error('VITE_TURSO_DB_URL not set!');

  return createClient({
    url: uri,
    authToken: token,
  });
}
