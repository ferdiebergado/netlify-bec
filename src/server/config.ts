import path from "path";

const rootDir = process.cwd();

const config = {
  paths: {
    public: path.join(rootDir, "public"),
    emTemplate: path.join(rootDir, "public", "em.xlsx"),
    beTemplate: "BLD-BE-001 Budget Estimate template.xlsx",
  },
  budgetEstimate: {
    cells: {
      program: "F4",
      output: "F5",
      outputIndicator: "F6",
      activity: "F7",
      activityIndicator: "F8",
      venue: "O13",
      totalPax: "H16",
    },
    rows: {
      boardLodging: {
        start: 17,
        end: 20,
        other: 24,
      },
      travelExpense: {
        regStart: 29,
        regEnd: 47,
        coStart: 48,
        coEnd: 50,
        other: 54,
      },
      honorarium: {
        start: 58,
        end: 59,
      },
      suppliesContingency: {
        start: 60,
        end: 61,
      },
    },
    lang: {
      en: {
        prefixes: {
          boardLodging: "Board and Lodging of ",
          travel: "Travel Expenses of ",
          honorarium: "Honorarium of ",
        },
      },
    },
  },
  sheets: {
    be: "BE-001",
    em: "Expenditure Form",
  },
};

export default config;
