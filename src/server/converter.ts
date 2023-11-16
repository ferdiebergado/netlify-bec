import BudgetEstimate from './budget_estimate';
import ExpenditureMatrix from './expenditure_matrix';
import fs from 'fs/promises';
import config from './config';
import { BUDGET_ESTIMATE, EXPENDITURE_MATRIX } from './constants';

export default async function convert(buffer: Buffer): Promise<ArrayBuffer> {
  const be = new BudgetEstimate(buffer, BUDGET_ESTIMATE.SHEET_NAME);
  await be.load();

  const emBuff = await fs.readFile(config.paths.emTemplate);
  const em = new ExpenditureMatrix(emBuff, EXPENDITURE_MATRIX.SHEET_NAME);
  await em.load();
  em.apply(be);
  const outBuff = await em.save();

  return outBuff;
}
