import { EXPENSE_GROUP, GAA_OBJECT, MANNER_OF_RELEASE } from './constants';

/**
 * Represents an expense group type based on the constants provided in the application.
 */
type ExpenseGroup = (typeof EXPENSE_GROUP)[keyof typeof EXPENSE_GROUP];

/**
 * Represents a GAA object type based on the constants provided in the application.
 */
type GAAObject = (typeof GAA_OBJECT)[keyof typeof GAA_OBJECT];

/**
 * Represents a manner of release type based on the constants provided in the application.
 */
type MannerOfRelease =
  (typeof MANNER_OF_RELEASE)[keyof typeof MANNER_OF_RELEASE];

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
  ppmp?: boolean;
  appSupplies?: boolean;
  appTicket?: boolean;
  mannerOfRelease: MannerOfRelease;
  [key: string]: any;
};

/**
 * Represents an activity, including program details, output, venue, etc., along with associated expense items.
 */
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

export { ExpenseGroup, GAAObject, MannerOfRelease, ExpenseItem, Activity };
