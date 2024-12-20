import type { Config } from './types/globals.js';

export const config: Readonly<Config> = {
  paths: {
    public: '/public',
    data: '/data',
    emTemplate: 'templates/em.xls',
    beTemplate: 'templates/BLD-BE-001 Budget Estimate template.xlsx',
  },
};
