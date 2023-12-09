import { Request, Response, NextFunction } from 'express';

export default function requestLogger(
  req: Request,
  _res: Response,
  next: NextFunction,
) {
  const { method, url, headers, query, body } = req;
  // eslint-disable-next-line no-console
  console.log(new Date().toString(), `${method} ${url}`, {
    headers,
    query,
    body,
  });
  next();
}
