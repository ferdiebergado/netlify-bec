import BudgetEstimate from './budget_estimate';
import ExpenditureMatrix from './expenditure_matrix';
import fs from 'fs/promises';
import config from './config';

export default async function convert(buffer: Buffer): Promise<ArrayBuffer> {
  const be = new BudgetEstimate(buffer);
  await be.load();

  const emBuff = await fs.readFile(config.paths.emTemplate);
  const em = new ExpenditureMatrix(emBuff);
  await em.load();
  em.apply(be);
  const outBuff = await em.save();

  return outBuff;
}
