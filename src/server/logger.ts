import { Request, Response, NextFunction } from 'express';
import executeQuery from './database';

export default async function requestLogger(
  req: Request,
  _res: Response,
  next: NextFunction,
) {
  Promise.resolve()
    .then(async () => {
      const { method, url, query } = req;
      const userAgent = req.get('User-Agent');
      const clientIp =
        req.ip ||
        (req.headers['x-forwarded-for'] || '').split(',').pop().trim() ||
        req.socket.remoteAddress;

      const createHitQuery = {
        sql: 'INSERT INTO hits (timestamp,method,url,query,ip,user_agent) VALUES (?,?,?,?,?,?);',
        args: [
          new Date().toISOString(),
          method,
          url,
          JSON.stringify(query),
          clientIp,
          userAgent,
        ],
      };

      const { lastInsertRowid } = await executeQuery(createHitQuery);

      req.insertedId = lastInsertRowid;

      next();
    })
    .catch(next);
}
