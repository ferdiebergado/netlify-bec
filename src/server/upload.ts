import { Request } from 'express';
import multer from 'multer';
import { EXCEL_MIMETYPE, MAX_FILESIZE, MAX_UPLOADS } from './constants';

function fileFilter(
  _req: Request,
  file: Express.Multer.File,
  cb: multer.FileFilterCallback,
): void {
  if (file.mimetype !== EXCEL_MIMETYPE) return cb(new Error('Wrong file type'));

  return cb(null, true);
}

const storage = multer.memoryStorage();
const upload = multer({
  storage,
  fileFilter,
  limits: { fileSize: MAX_FILESIZE },
}).array('excelFile', MAX_UPLOADS);

export default upload;
