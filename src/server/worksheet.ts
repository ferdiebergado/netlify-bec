import Excel from 'exceljs';

class Worksheet {
  static LOAD_ERROR_MSG = 'Worksheet not set.  Call the load() method first.';

  wb: Excel.Workbook;

  ws?: Excel.Worksheet;

  constructor(
    protected xls: ArrayBuffer,
    protected sheet: string,
  ) {
    this.wb = new Excel.Workbook();
  }

  async load() {
    await this.wb.xlsx.load(this.xls);

    this.ws = this.wb.getWorksheet(this.sheet);
  }
}

export default Worksheet;
