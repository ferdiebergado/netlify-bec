import path from 'path';

const rootDir = process.cwd();

const config = {
  paths: {
    public: path.join(rootDir, 'public'),
    emTemplate: path.join(rootDir, 'data', 'em.xlsx'),
    beTemplate: 'BLD-BE-001 Budget Estimate template.xlsx',
  },
  db: {
    uri: process.env.MONGODB_URI || 'mongodb://localhost:27017',
    databaseName: process.env.MONGODB_DBNAME || 'bec',
  },
};

export default config;
