// eslint-disable-next-line import/no-extraneous-dependencies
import { InStatement, ResultSet } from '@libsql/client/.';
import useTurso from './turso';

/**
 * Executes a SQL query using the Turso database client.
 *
 * @param {InStatement} query - The SQL query to be executed.
 * @returns {Promise<ResultSet>} - A promise that resolves to the result of the query execution.
 *
 * @throws {Error} If there is an issue with the query execution.
 *
 * @example
 * const query = {sql: 'SELECT * FROM tableName WHERE column = ?', args: ['value'])};
 * try {
 *   const result = await executeQuery(query);
 *   console.log(result);
 * } catch (error) {
 *   console.error('Error executing query:', error.message);
 * }
 */
export default async function executeQuery(
  query: InStatement,
): Promise<ResultSet> {
  /**
   * Uses the Turso database client to execute a query.
   *
   * @returns {Object} - A Turso database client instance.
   *
   * @throws {Error} If there is an issue with creating or obtaining the Turso client.
   *
   * @example
   * const client = useTurso();
   */
  const client = useTurso();

  try {
    const result = await client.execute(query);
    return result;
  } catch (error: any) {
    throw new Error(`Error executing query: ${error.message}`);
  }
}
