// eslint-disable-next-line @typescript-eslint/no-redeclare
import type { NextFunction, Request, Response } from 'express';
import express, { Router } from 'express';
import { BASE_URL, CONVERT_URL, EXCEL_MIMETYPE } from './constants';
import convert from './converter';
import { createTimestamp } from './utils';
import logger from './logger';
import upload from './upload';
import errorHandler from './error_handler';
import { updateDocument } from './database';

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
      await updateDocument(
        'requests',
        { _id: req.insertedId },
        { $set: { uploads } },
      );

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

router.use(logger);
router.post(CONVERT_URL, upload, handleConvert);
router.use(errorHandler);

const server = express();

server.disable('x-powered-by');
server.use(BASE_URL, router);

export default server;
