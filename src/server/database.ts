// eslint-disable-next-line import/no-extraneous-dependencies
import { InStatement } from '@libsql/client/.';
import useTurso from './turso';

export default async function executeQuery(query: InStatement) {
  const client = useTurso();

  const result = await client.execute(query);
  return result;
}
