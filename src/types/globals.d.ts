import { Worksheet } from 'exceljs';
import { ReleaseManner, GAAObject, ExpenseGroup } from '../constants';

/**
 * Represents an activity, including program details, output, venue, etc., along with associated expense items.
 *
 * @interface Activity
 */
interface Activity {
  info: ActivityInfo;
  tevPSF: ExpenseItem[];
  expenseItems: ExpenseItem[];
}

/**
 * Represents an activity info including program, output, venue, etc.
 *
 * @interface ActivityInfo
 */
interface ActivityInfo {
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
}

/**
 * Represents an expense item, including its details such as expense group, GAA object, etc.
 *
 * @interface ExpenseItem
 */
interface ExpenseItem {
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
}

/**
 * Options for managing expenses.
 */
interface ExpenseOptions {
  /**
   * A prefix to identify expense records.
   */
  prefix: string;

  /**
   * The manner in which the release is handled.
   */
  releaseManner: ReleaseManner;

  /**
   * Indicates if a PPMP (Project Procurement Management Plan) is associated.
   * Optional.
   */
  hasPPMP?: boolean;

  /**
   * The venue associated with the expense. Optional.
   */
  venue?: string;
}

/**
 * Represents buffer types, either a Node.js Buffer or a generic ArrayBuffer.
 */
type Buffers = Buffer | ArrayBuffer;

/**
 * Represents an Excel file with a filename and its content as a buffer.
 */
interface ExcelFile {
  /**
   * The name of the Excel file.
   */
  filename: string;

  /**
   * The content of the file, represented as a buffer.
   */
  buffer: Buffers;
}

/**
 * Utility type for making all properties of a type optional, recursively.
 */
type DeepPartial<T> = T extends object
  ? {
      [P in keyof T]?: DeepPartial<T[P]>;
    }
  : T;

/**
 * Details of a backend parsing error.
 *
 * @interface BEParseErrDetail
 */
interface BEParseErrDetail {
  /**
   * The activity related to the error. Optional.
   */
  activity?: string;

  /**
   * The file where the error occurred.
   */
  file: string;

  /**
   * The sheet in the file where the error occurred. Optional.
   */
  sheet?: string;
}

/**
 * Paths used in the configuration.
 *
 * @interface Paths
 */
interface Paths {
  /**
   * Path to the data directory.
   */
  data: string;

  /**
   * Path to the Budget Estimate template.
   */
  beTemplate: string;
}

/**
 * Configuration object.
 *
 * @interface Config
 */
interface Config {
  /**
   * The paths used in the configuration.
   */
  paths: Paths;
}

/**
 * Configuration for processing a sheet in an Excel file.
 *
 * @interface SheetConfig
 */
interface SheetConfig {
  /**
   * The index of the starting row in the sheet.
   */
  startRowIndex: number;

  /**
   * The index of the starting column in the sheet.
   */
  startColIndex: number;

  /**
   * The number of rows to process.
   */
  numRows: number;

  /**
   * Options for managing expenses in the sheet.
   */
  options: ExpenseOptions;
}

/**
 * Map for copying rows between indices in a sheet.
 *
 * @interface RowCopyMap
 */
interface RowCopyMap {
  /**
   * The index of the source row.
   */
  srcRowIndex: number;

  /**
   * The number of rows to copy.
   */
  numRows: number;
}

/**
 * Represents the context of an expense item row.
 *
 * @interface ExpenseItemRowContext
 */
interface ExpenseItemRowContext {
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
}

/**
 * Represents the context of an Activity.
 *
 * @interface ActivityContext
 */
interface ActivityRowMap {
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
}

interface OverheadTotalRowMap {
  rowsToAdd: number;
  expenseItemsCount?: number;
}

interface ExpenditureFile {
  programTitle: string | undefined;
  buffer: ArrayBuffer;
}
