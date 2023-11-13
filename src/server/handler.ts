import BudgetEstimate from './budget_estimate';
import ExpenditureMatrix from './expenditure_matrix';
import fs from 'fs/promises';
import config from './config';
import { BUDGET_ESTIMATE, EXPENDITURE_MATRIX } from './constants';

class Handler {
  async dispatch(req: Request): Promise<Response> {
    const beBuff = await req.arrayBuffer();
    const be = new BudgetEstimate(beBuff, BUDGET_ESTIMATE.SHEET_NAME);
    await be.load();
    const emBuff = await fs.readFile(config.paths.emTemplate);
    const em = new ExpenditureMatrix(emBuff, EXPENDITURE_MATRIX.SHEET_NAME);
    await em.load();
    em.apply(be);
    const outBuff = await em.save();

    // Create a Blob from the buffer
    const blob = new Blob([outBuff]);

    const res = new Response(blob);
    const filename = `em-${new Date().getTime()}.xlsx`;

    res.headers.set(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    res.headers.set('Content-Disposition', `attachment; filename=${filename}`);

    return res;
  }
}

export default Handler;
