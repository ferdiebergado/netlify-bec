// eslint-disable-next-line @typescript-eslint/no-redeclare
import type { NextFunction, Request, Response } from 'express';
import express, { Router } from 'express';
import { BASE_URL, CONVERT_URL, EXCEL_MIMETYPE } from './constants';
import convert from './converter';
import { createTimestamp } from './utils';
import logger from './logger';
import upload from './upload';
import errorHandler from './error_handler';
import executeQuery from './database';

/**
 * Handles the conversion of files to Excel format and sends the converted file as a response.
 *
 * @param req - The Express request object.
 * @param res - The Express response object.
 * @param next - The next middleware function.
 * @returns A Promise that resolves once the conversion is complete.
 */
async function handleConvert(
  req: Request,
  res: Response,
  next: NextFunction,
): Promise<void> {
  Promise.resolve()
    .then(async () => {
      if (!req.files) throw new Error('File is required.');

      const files = req.files as Express.Multer.File[];
      // const buffers: Buffer[] = files.map(file => file.buffer);

      const uploads = files.map(file => file.originalname);
      const updateHitQuery = {
        sql: 'UPDATE hits SET files = ? WHERE hit_id = ?;',
        args: [uploads.toString(), req.insertedId!],
      };

      await executeQuery(updateHitQuery);

      const outBuff = await convert(files);
      const filename = `em-${createTimestamp()}.xlsx`;

      res
        .header({
          'Content-Type': EXCEL_MIMETYPE,
          'Content-Disposition': `attachment; filename=${filename}`,
          'Content-Length': outBuff.byteLength.toString(),
        })
        .end(outBuff);
    })
    .catch(next);
}

const router = Router();

// Logger middleware
router.use(logger);

// Route for handling file conversion
router.post(CONVERT_URL, upload, handleConvert);

// Error handling middleware
router.use(errorHandler);

const server = express();

server.disable('x-powered-by');
server.use(BASE_URL, router);

export default server;
