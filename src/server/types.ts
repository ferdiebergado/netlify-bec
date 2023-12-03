import { CellValue, DataValidation } from 'exceljs';
import { EXPENSE_GROUP, GAA_OBJECT, MANNER_OF_RELEASE } from './constants';

type ExpenseGroup = (typeof EXPENSE_GROUP)[keyof typeof EXPENSE_GROUP];

type GAAObject = (typeof GAA_OBJECT)[keyof typeof GAA_OBJECT];

type MannerOfRelease =
  (typeof MANNER_OF_RELEASE)[keyof typeof MANNER_OF_RELEASE];

type ExpenseItem = {
  expenseGroup: ExpenseGroup;
  gaaObject: GAAObject;
  expenseItem: string;
  quantity: number;
  freq?: number;
  unitCost: number;
  tevLocation?: string;
  ppmp?: boolean;
  appSupplies?: boolean;
  appTicket?: boolean;
  mannerOfRelease: MannerOfRelease;
  [key: string]: any;
};

type Activity = {
  program: string;
  output: string;
  outputIndicator: string;
  activityTitle: string;
  activityIndicator: string;
  month: number;
  venue: string;
  totalPax: number;
  outputPhysicalTarget: number;
  activityPhysicalTarget: number;
  expenseItems: ExpenseItem[];
};

type CellData = {
  cell: string;
  value: CellValue;
  dataValidation?: DataValidation;
};

export {
  ExpenseGroup,
  GAAObject,
  MannerOfRelease,
  ExpenseItem,
  Activity,
  CellData,
};
