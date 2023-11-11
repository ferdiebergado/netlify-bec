import type { Context } from "@netlify/functions";
import Handler from "../../src/server/handler";

export default async (req: Request, _context: Context) => {
  if (req.method == "post") {
    return await new Handler().dispatch(req);
  }
};
