import type { Request, Response } from 'express';
import express, { Router } from 'express';
import { BASE_URL, CONVERT_URL, EXCEL_MIMETYPE } from './constants';
import convert from './converter';
import { asyncMiddlewareWrapper, createTimestamp } from './utils';
import upload from './upload';
import errorHandler from './error_handler';

/**
 * Handles the conversion of files to Excel format and sends the converted file as a response.
 *
 * @param req The Express request object.
 * @param res The Express response object.
 * @param next The next middleware function.
 *
 * @returns A Promise that resolves once the conversion is complete.
 */
async function conversionHandler(req: Request, res: Response): Promise<void> {
  if (!req.files) throw new Error('File is required.');

  const files = req.files as Express.Multer.File[];

  // if (process.env.NODE_ENV === 'production') {
  //   const uploads = files.map(file => file.originalname);
  //   const updateHitQuery = {
  //     sql: 'UPDATE hits SET files = ? WHERE hit_id = ?;',
  //     args: [uploads.toString(), req.insertedId!],
  //   };

  //   await executeQuery(updateHitQuery);
  // }

  const outBuff = await convert(files);
  const filename = `em-${createTimestamp()}.xlsx`;

  res
    .header({
      'Content-Type': EXCEL_MIMETYPE,
      'Content-Disposition': `attachment; filename=${filename}`,
      'Content-Length': outBuff.byteLength.toString(),
    })
    .end(outBuff);
}

const router = Router();

// Logger middleware
// router.use(asyncMiddlewareWrapper(logger));

// Route for handling file conversion
// eslint-disable-next-line @typescript-eslint/no-misused-promises
router.post(CONVERT_URL, upload, asyncMiddlewareWrapper(conversionHandler));

// Error handling middleware
router.use(errorHandler);

const server = express();

server.disable('x-powered-by');
server.use(BASE_URL, router);

export default server;
