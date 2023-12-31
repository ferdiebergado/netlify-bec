import type { Worksheet } from 'exceljs';
import Excel from 'exceljs';

/**
 * Represents an abstract workbook with basic functionality for working with Excel files using the exceljs library.
 *
 * @abstract
 * @class Workbook
 * @template T - The type of the concrete Workbook class that extends this abstract class.
 */
export abstract class Workbook<T extends Workbook<T>> {
  /**
   * The instance of the Excel Workbook used for operations.
   *
   * @protected
   * @type {Excel.Workbook}
   */
  protected wb: Excel.Workbook = new Excel.Workbook();

  /**
   * The instance of the Excel Worksheet (optional) within the workbook.
   *
   * @protected
   * @type {Worksheet|undefined}
   */
  protected ws?: Worksheet;

  /**
   * The source of the workbook, which can be either a file path (string) or a Buffer containing workbook data.
   *
   * @protected
   * @type {string|Buffer|undefined}
   */
  protected source?: string | Buffer;

  /**
   * Creates an instance of the concrete Workbook class.
   *
   * @protected
   * @abstract
   * @returns {T} - The created instance of the concrete Workbook class.
   */
  protected abstract createInstance(): T;

  /**
   * Asynchronously initializes the workbook with data from the provided source.
   *
   * @public
   * @param {string|Buffer} source - The source of the workbook data, either a file path (string) or a Buffer.
   * @returns {Promise<void>} - A Promise that resolves when the initialization is complete.
   */
  async initializeAsync(source: string | Buffer): Promise<void> {
    if (typeof source === 'string') {
      await this.wb.xlsx.readFile(source);
    } else if (source instanceof Buffer) {
      await this.wb.xlsx.load(source);
    }

    this.source = source;
  }

  /**
   * Asynchronously creates an instance of the concrete Workbook class and initializes it with data from the provided source.
   *
   * @public
   * @static
   * @param {string|Buffer} source - The source of the workbook data, either a file path (string) or a Buffer.
   * @returns {Promise<T>} - A Promise that resolves with the created and initialized instance of the concrete Workbook class.
   */
  static async createAsync<T extends Workbook<T>>(
    source: string | Buffer,
  ): Promise<T> {
    const instance = new (this as unknown as { new (): T })();
    await instance.initializeAsync(source);
    return instance.createInstance();
  }

  /**
   * Fetches the active worksheet.
   *
   * @throws {Error} Will throw an error if the active sheet is not set
   *
   * @returns {Worksheet} The current worksheet or an error
   */
  getActiveSheet(): Worksheet {
    if (!this.ws) throw new Error('Active worksheet not set!');

    return this.ws;
  }

  /**
   * Sets the active worksheet to the specified index or name.
   *
   * @param indexOrName {number|string} The index or name of the worksheet
   */
  setActiveSheet(indexOrName: number | string): void {
    this.ws = this.wb.getWorksheet(indexOrName);
  }
}
