import { Request, Response, NextFunction } from 'express';

export default function errorHandler(
  err: Error,
  _req: Request,
  res: Response,
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  _next: NextFunction,
) {
  // eslint-disable-next-line no-console
  console.error(err.stack);
  res.status(500).send({ error: 'Conversion failed!' });
}
