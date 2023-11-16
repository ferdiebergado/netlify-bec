import path from 'path';

const rootDir = process.cwd();

const config = {
  paths: {
    public: path.join(rootDir, 'public'),
    emTemplate: path.join(rootDir, 'data', 'em.xlsx'),
    beTemplate: 'BLD-BE-001 Budget Estimate template.xlsx',
  },
};

export default config;
