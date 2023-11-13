import type { Context } from '@netlify/functions';
import Handler from '../../src/server/handler';

// eslint-disable-next-line @typescript-eslint/no-unused-vars
export default async (req: Request, _context: Context) => {
  if (req.method === 'POST') {
    const response = await new Handler().dispatch(req);
    return response;
  }
};
