// eslint-disable-next-line @typescript-eslint/no-redeclare
import type { NextFunction, Request, Response } from 'express';
import express, { Router } from 'express';
import multer from 'multer';
import { EXCEL_MIMETYPE, MAX_FILESIZE, MAX_UPLOADS } from './constants';
import convert from './converter';
import { createTimestamp } from './utils';

const router = Router();
const storage = multer.memoryStorage();

function errorHandler(
  err: Error,
  _req: Request,
  res: Response,
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  _next: NextFunction,
) {
  console.error(err.stack);
  res.status(500).send({ error: 'Conversion failed!' });
}

function fileFilter(
  _req: Request,
  file: Express.Multer.File,
  cb: multer.FileFilterCallback,
): void {
  if (file.mimetype !== EXCEL_MIMETYPE) return cb(new Error('Wrong file type'));

  if (file.size > MAX_FILESIZE) return cb(new Error('File too large'));

  cb(null, true);
}

async function handleConvert(
  req: Request,
  res: Response,
  next: NextFunction,
): Promise<void> {
  Promise.resolve()
    .then(async () => {
      if (!req.files) throw new Error('File is required.');

      const files = req.files as Express.Multer.File[];
      const buffers: Buffer[] = files.map(file => file.buffer);
      const outBuff = await convert(buffers);
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

const upload = multer({ storage, fileFilter }).array('excelFile', MAX_UPLOADS);

router.post('/convert', upload, handleConvert);
router.use(errorHandler);

const server = express();

server.disable('x-powered-by');
server.use('/api/', router);

export default server;
