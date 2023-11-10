import type { Context } from "@netlify/functions";
import excel from "exceljs";
import path from "path";
import fs from "fs";

const { Workbook } = excel;

export default async (req: Request, context: Context) => {
  const wb = new Workbook();
  const be = path.join(process.cwd(), "public", "be.xlsx");
  // await wb.xlsx.readFile(path.join(__dirname, "be.xlsx"));
  await wb.xlsx.readFile(be);
  const ws = wb.getWorksheet("BAE-BE-001");

  const prog = ws?.getCell("F3").text;
  const xlsx = path.join(process.cwd(), "be (1).xlsx");
  // await wb.xlsx.writeFile(path.join(process.cwd(), "be (1).xlsx"));

  // const file = fs.readFileSync(xlsx)
  const res = new Response();

  res.headers.set("Content-Length", fs.statSync(xlsx).size.toString());
  res.headers.set(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  res.headers.set("Content-Disposition", "attachment; filename=bebe.xlsx");
  // res.write(file, 'binary');

  // const res = await fetch(xlsx);
  // const blob = await res.blob();
  // const url = URL.createObjectURL(blob);

  return res;
};
