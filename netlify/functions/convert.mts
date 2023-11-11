import type { Context } from "@netlify/functions";
import Handler from "../../src/server/handler";

export default async (req: Request, _context: Context) => {
  return await new Handler().dispatch(req);
};
