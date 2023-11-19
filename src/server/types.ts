import { CellValue, DataValidation } from 'exceljs';
import { EXPENSE_GROUP, GAA_OBJECT, MANNER_OF_RELEASE } from './constants';

type ExpenseGroup = (typeof EXPENSE_GROUP)[keyof typeof EXPENSE_GROUP];

type GAAObject = (typeof GAA_OBJECT)[keyof typeof GAA_OBJECT];

type MannerOfRelease =
  (typeof MANNER_OF_RELEASE)[keyof typeof MANNER_OF_RELEASE];

type YesNo = 'Y' | 'N';

type ExpenseItem = {
  expenseGroup?: ExpenseGroup;
  gaaObject?: GAAObject;
  expenseItem: string;
  quantity: number;
  freq: number;
  unitCost: number;
  tevLocation: string;
  ppmp?: YesNo;
  appSupplies: YesNo;
  appTicket: YesNo;
  mannerOfRelease: MannerOfRelease;
};

type ActivityInfo = {
  program?: string;
  output?: string;
  outputIndicator?: string;
  activity?: string;
  activityIndicator?: string;
  month?: number;
  venue: string;
  totalPax: number;
};

type CellData = {
  cell: string;
  value: CellValue;
  dataValidation?: DataValidation;
};

export {
  YesNo,
  ExpenseGroup,
  GAAObject,
  MannerOfRelease,
  ExpenseItem,
  ActivityInfo,
  CellData,
};
