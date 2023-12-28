import { Request, Response, NextFunction } from 'express';
import executeQuery from './database';
import { InStatement, InValue } from '@libsql/client/.';

export default function requestLogger(
  req: Request,
  _res: Response,
  next: NextFunction,
) {
  const { method, url, query } = req;
  const userAgent: InValue = req.get('User-Agent') || 'Unknown User-Agent';
  const xForwardedForHeader = req.headers['x-forwarded-for'];

  let clientIp: InValue;

  if (typeof xForwardedForHeader === 'string') {
    // If it's a string, split it and get the last IP address
    const ipAddressArray = xForwardedForHeader.split(',');
    clientIp = ipAddressArray[ipAddressArray.length - 1].trim();
  } else if (Array.isArray(xForwardedForHeader)) {
    // If it's an array, assume it's an array of IP addresses and get the last one
    clientIp = xForwardedForHeader[xForwardedForHeader.length - 1].trim();
  } else {
    clientIp = req.socket.remoteAddress || 'unknown ip';
  }

  const createHitQuery: InStatement = {
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

  Promise.resolve()
    .then(async () => {
      const result = await executeQuery(createHitQuery);
      const { lastInsertRowid } = result;
      req.insertedId = lastInsertRowid;
    })
    .catch(next);
}
