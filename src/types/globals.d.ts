import { Worksheet } from 'exceljs';
import { ReleaseManner, GAAObject, ExpenseGroup } from '../constants';

/**
 * Represents an activity, including program details, output, venue, etc., along with associated expense items.
 */
type Activity = {
  info: ActivityInfo;
  tevPSF: ExpenseItem[];
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
  totalPax?: number;
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

type Buffers = Buffer | ArrayBuffer;

type ExcelFile = {
  filename: string;
  buffer: Buffers;
};

type DeepPartial<T> = T extends object
  ? {
      [P in keyof T]?: DeepPartial<T[P]>;
    }
  : T;

type BEParseErrDetail = {
  activity?: string;
  file: string;
  sheet?: string;
};

interface Paths {
  public: string;
  data: string;
  emTemplate: string;
  beTemplate: string;
}

interface Config {
  paths: Paths;
}

interface SheetConfig {
  startRowIndex: number;
  startColIndex: number;
  numRows: number;
  options: ExpenseOptions;
}

interface RowCopyMap {
  targetRowIndex: number;
  srcRowIndex: number;
  numRows: number;
}

/**
 * Represents the context of an expense item row.
 *
 * @interface ExpenseItemRowContext
 */
interface ExpenseItemRowContext {
  /**
   * The row number where the expense item row will be inserted.
   * @type {number}
   */
  targetRowIndex: number;

  /**
   * The expense item data.
   * @type {ExpenseItem}
   */
  expense: ExpenseItem;

  /**
   * The month index.
   * @type {number}
   */
  month: number;

  /**
   * Flag indicating if the activity being created is the very first activity. Default is `false`.
   * @type {boolean}
   */
  isFirstActivity: boolean;
}

/**
 * Represents the context of an Activity.
 *
 * @interface ActivityContext
 */
interface ActivityContext {
  /**
   * The activity being processed
   * @type {Activity}
   */
  activity: Activity;

  /**
   * The current row index
   * @type {number}
   */
  rowIndex: number;

  /**
   * Indicates if the activity is the first activity to be written to the active sheet.
   * @type {boolean}
   */
  isFirstActivity: boolean;
}
