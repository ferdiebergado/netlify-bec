import type { Request } from 'express';
import multer from 'multer';
import { EXCEL_MIMETYPE, MAX_FILESIZE, MAX_UPLOADS } from './constants';

/**
 * Middleware to validate if the uploaded file is an Excel file.
 *
 * @param _req The incoming Request object
 * @param file The uploaded file
 * @param cb The callback function to filter the file
 *
 * @returns void
 */
function fileFilter(
  _req: Request,
  file: Express.Multer.File,
  cb: multer.FileFilterCallback,
): void {
  if (file.mimetype !== EXCEL_MIMETYPE) return cb(new Error('Wrong file type'));

  return cb(null, true);
}

/**
 * The storage to be used to store the file uploads.
 *
 * @type multer.StorageEngine
 */
const storage: multer.StorageEngine = multer.memoryStorage();

/**
 * Middleware that will handle file uploads.
 */
const upload = multer({
  storage,
  fileFilter,
  limits: { fileSize: MAX_FILESIZE },
}).array('excelFile', MAX_UPLOADS);

export default upload;
