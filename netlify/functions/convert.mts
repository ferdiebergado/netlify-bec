import type { Context } from "@netlify/functions";
import excel from "exceljs";
import path from "path";
import fs from "fs";

const { Workbook } = excel;

export default async (req: Request, context: Context) => {
  const wb = new Workbook();
  // const be = path.join(process.cwd(), "public", "be.xlsx");
  // await wb.xlsx.readFile(be);

  // const xlsx = path.join(process.cwd(), "be (1).xlsx");
  // await wb.xlsx.writeFile(path.join(process.cwd(), "be (1).xlsx"));

  // const file = fs.readFileSync(xlsx)

  const buffer = await req.arrayBuffer();
  await wb.xlsx.load(buffer);
  const ws = wb.getWorksheet("BAE-BE-001");

  if (ws) {
    // prog = ws?.getCell("F3").text;

    ws.getCell("F3").value = "test";
  }

  const buf2 = await wb.xlsx.writeBuffer();

  // Create a Blob from the buffer
  const blob = new Blob([buf2]);

  const res = new Response(blob);
  const filename = `em-${new Date().getTime()}.xlsx`;

  // res.headers.set("Content-Length", fs.statSync(xlsx).size.toString());
  res.headers.set(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  res.headers.set("Content-Disposition", `attachment; filename=${filename}`);

  return res;
};
