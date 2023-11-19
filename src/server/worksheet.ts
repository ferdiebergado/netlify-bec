import Excel from 'exceljs';

export default class Worksheet {
  static readonly LOAD_ERROR_MSG =
    'Worksheet not set.  Call the load() method first.';

  protected wb: Excel.Workbook;

  protected ws?: Excel.Worksheet;

  constructor(
    protected readonly xls: ArrayBuffer,
    protected readonly sheet: string,
  ) {
    this.wb = new Excel.Workbook();
  }

  async load() {
    await this.wb.xlsx.load(this.xls);

    this.ws = this.wb.getWorksheet(this.sheet);
  }
}
