import config from './config';
import type { ExcelFile } from '../types/globals';
import { ExpenditureMatrix } from './expenditureMatrix';

/**
 * Converts uploaded files into an ArrayBuffer representing the expenditure matrix.
 *
 * @param {Express.Multer.File[]} files The array of uploaded files.
 *
 * @returns {Promise<ArrayBuffer>} A promise that resolves to the ArrayBuffer of the expenditure matrix.
 */
export default async function convert(
  files: Express.Multer.File[],
): Promise<ArrayBuffer> {
  const emTemplate = config.paths.emTemplate;
  const em = await ExpenditureMatrix.createAsync<ExpenditureMatrix>(emTemplate);

  const excelFiles: ExcelFile[] = files.map(file => {
    const { originalname, buffer } = file;
    return { filename: originalname, buffer };
  });

  const buffer = await em.convert(excelFiles);

  return buffer;
}
