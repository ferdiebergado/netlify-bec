import Excel from 'exceljs';

class Worksheet {
  wb: Excel.Workbook;

  ws?: Excel.Worksheet;

  constructor(
    protected xls: ArrayBuffer,
    protected sheet: string,
  ) {
    this.wb = new Excel.Workbook();
    this.ws = undefined;
  }

  async load() {
    await this.wb.xlsx.load(this.xls);

    this.ws = this.wb.getWorksheet(this.sheet);
  }
}

export default Worksheet;
