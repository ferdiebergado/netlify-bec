import config from './config';
import type { ExcelFile } from './types/globals';
import { ExpenditureMatrix } from './expenditureMatrix';

/**
 * Converts uploaded files into an ArrayBuffer representing the expenditure matrix.
 *
 * @param {FileList} files The array of uploaded files.
 *
 * @returns {Promise<ArrayBuffer>} A promise that resolves to the ArrayBuffer of the expenditure matrix.
 */
export default async function convert(files: FileList): Promise<ArrayBuffer> {
  const emTemplate = config.paths.emTemplate;
  const res = await fetch(emTemplate);
  const arrayBuffer = await res.arrayBuffer();
  const expenditureMatrix =
    await ExpenditureMatrix.createAsync<ExpenditureMatrix>(arrayBuffer);

  const excelFiles: ExcelFile[] = [];

  await Promise.allSettled(
    [...files].map(async file => {
      excelFiles.push({
        filename: file.name,
        buffer: await file.arrayBuffer(),
      });
    }),
  );

  const buffer = await expenditureMatrix.fromBudgetEstimates(excelFiles);

  return buffer;
}
