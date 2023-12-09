import { Request, Response, NextFunction } from 'express';
import { insertDocument } from './database';

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

      const document = {
        timestamp: new Date().getTime(),
        method,
        url,
        query,
        ip: clientIp,
        userAgent,
      };

      req.insertedId = await insertDocument(document, 'requests');

      next();
    })
    .catch(next);
}
