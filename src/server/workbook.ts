import type { Worksheet } from 'exceljs';
import Excel from 'exceljs';

export abstract class Workbook<T extends Workbook<T>> {
  protected wb: Excel.Workbook = new Excel.Workbook();
  protected ws?: Worksheet;
  protected source?: string | Buffer;

  constructor() {}

  protected abstract createInstance(): T;

  async initializeAsync(source: string | Buffer): Promise<void> {
    if (typeof source === 'string') {
      await this.wb.xlsx.readFile(source);
    } else if (source instanceof Buffer) {
      await this.wb.xlsx.load(source);
    }

    this.source = source;
  }

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
