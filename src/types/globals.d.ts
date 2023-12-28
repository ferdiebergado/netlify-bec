import { ReleaseManner, GAAObject, ExpenseGroup } from '../server/constants';

/**
 * Represents an activity, including program details, output, venue, etc., along with associated expense items.
 */
type Activity = {
  info: ActivityInfo;
  expenseItems: ExpenseItem[];
};

/**
 * Represents an activity info including program, output, venue, etc.
 */
type ActivityInfo = {
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
};

/**
 * Represents an expense item, including its details such as expense group, GAA object, etc.
 */
type ExpenseItem = {
  expenseGroup: ExpenseGroup;
  gaaObject: GAAObject;
  expenseItem: string;
  quantity: number;
  freq?: number;
  unitCost: number;
  tevLocation?: string;
  hasPPMP?: boolean;
  hasAPPSupplies?: boolean;
  hasAPPTicket?: boolean;
  releaseManner: ReleaseManner;
};

type ExpenseOptions = {
  prefix: string;
  releaseManner: ReleaseManner;
  hasPPMP?: boolean;
  venue?: string;
};

type ExcelFile = {
  filename: string;
  buffer: Buffer;
};
