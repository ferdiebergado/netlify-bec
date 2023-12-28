import path from 'node:path';

const rootDir = process.cwd();

const config = {
  paths: {
    public: path.join(rootDir, 'public'),
    data: path.join(rootDir, 'data'),
    emTemplate: path.join(rootDir, 'data', 'em.xlsx'),
    beTemplate: 'BLD-BE-001 Budget Estimate template.xlsx',
  },
  db: {
    uri: process.env.VITE_TURSO_DB_URL,
    token: process.env.VITE_TURSO_DB_AUTH_TOKEN,
  },
};

export default config;
