import type { BEParseErrDetail } from './types/globals.js';

export class BudgetEstimateParseError extends Error {
  constructor(
    message: string,
    public details: BEParseErrDetail,
  ) {
    super(message);
    this.name = this.constructor.name;

    if (Error.captureStackTrace) {
      Error.captureStackTrace(this, this.constructor);
    }
  }
}
