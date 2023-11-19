import serverless from 'serverless-http';
import server from '../../src/server/server';

export const handler = serverless(server, { binary: true });
