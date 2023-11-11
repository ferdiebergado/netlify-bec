import Excel from "exceljs";

class Worksheet {
  xls: ArrayBuffer;

  sheet: string;

  wb: Excel.Workbook;

  ws?: Excel.Worksheet;

  constructor(xls: ArrayBuffer, sheet: string) {
    this.xls = xls;
    this.sheet = sheet;
    this.wb = new Excel.Workbook();
    this.ws = undefined;
  }

  async load() {
    await this.wb.xlsx.load(this.xls);

    this.ws = this.wb.getWorksheet(this.sheet);
  }
}

export default Worksheet;
