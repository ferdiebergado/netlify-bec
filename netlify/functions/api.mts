// eslint-disable-next-line @typescript-eslint/no-redeclare
import express, { Request, Response, Router } from 'express';
import serverless from 'serverless-http';
import multer from 'multer';
import { EXCEL_MIMETYPE } from '../../src/server/constants';
import convert from '../../src/server/converter';
import { timestamp } from '../../src/server/utils';

const api = express();
const router = Router();
const storage = multer.memoryStorage();

function fileFilter(
  _req: Request,
  file: Express.Multer.File,
  cb: multer.FileFilterCallback,
) {
  if (file.mimetype !== EXCEL_MIMETYPE) return cb(new Error('Wrong file type'));

  cb(null, true);
}

async function handleConvert(req: Request, res: Response) {
  const { file } = req;

  if (!file) throw new Error('File is required.');

  const beBuff = file.buffer;
  const outBuff = await convert(beBuff);
  const filename = `em-${timestamp()}.xlsx`;

  res
    .header({
      'Content-Type': EXCEL_MIMETYPE,
      'Content-Disposition': `attachment; filename=${filename}`,
      'Content-Length': outBuff.byteLength.toString(),
    })
    .end(outBuff);
}

const upload = multer({ storage, fileFilter }).single('excelFile');

router.post('/convert', upload, handleConvert);

api.use('/api/', router);

export const handler = serverless(api, { binary: true });
