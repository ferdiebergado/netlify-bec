import path from 'path';
import { BudgetEstimate } from '../../src/server/budgetEstimate';
import config from '../../src/server/config';
import type { ActivityInfo, ExpenseItem } from '../../src/types/globals';

describe('BudgetEstimate class', () => {
  const budgetEstimateFile = path.join(config.paths.data, 'be_test.xlsx');
  let budgetEstimate: BudgetEstimate;

  beforeEach(async () => {
    budgetEstimate = await BudgetEstimate.createAsync(budgetEstimateFile);
    budgetEstimate.setActiveSheet(1);
  });

  it('should correctly parse activity information', () => {
    const expected: ActivityInfo = {
      program: 'Cosmic Learning System',
      output: 'Oriented galaxy heads',
      outputIndicator: 'No. of galaxy heads oriented',
      activityTitle: 'Orientation of Galaxy Heads on Cosmic Education System',
      activityIndicator: 'No. of orientations conducted',
      month: 7,
      venue: 'BUTUAN',
      totalPax: 85,
      outputPhysicalTarget: 150,
      activityPhysicalTarget: 1,
    };

    const result = budgetEstimate.getActivityInfo();

    expect(result).toEqual(expected);
  });

  describe('Board and Lodging expenses', () => {
    it('should correctly get board and lodging expenses of participants', () => {
      const expected: Partial<ExpenseItem> = {
        expenseItem: 'Board and Lodging of Participants',
        quantity: 41,
        freq: 5,
        unitCost: 1500,
      };

      const result = budgetEstimate.getBoardAndLodging();
      const actual = result[0];

      expect(actual.expenseItem).toEqual(expected.expenseItem);
      expect(actual.quantity).toEqual(expected.quantity);
      expect(actual.freq).toEqual(expected.freq);
      expect(actual.unitCost).toEqual(expected.unitCost);
    });

    it('should correctly get board and lodging expenses of resource persons', () => {
      const expected: Partial<ExpenseItem> = {
        expenseItem: 'Board and Lodging of Resource Persons',
        quantity: 8,
        freq: 5,
        unitCost: 2000,
      };

      const result = budgetEstimate.getBoardAndLodging();
      const actual = result[1];

      expect(actual.expenseItem).toEqual(expected.expenseItem);
      expect(actual.quantity).toEqual(expected.quantity);
      expect(actual.freq).toEqual(expected.freq);
      expect(actual.unitCost).toEqual(expected.unitCost);
    });

    it('should correctly get board and lodging expenses of technical experts', () => {
      const expected: Partial<ExpenseItem> = {
        expenseItem:
          'Board and Lodging of Technical Experts ( Editors/Validators/Illustrators)',
        quantity: 4,
        freq: 5,
        unitCost: 2000,
      };

      const result = budgetEstimate.getBoardAndLodging();
      const actual = result[2];

      expect(actual.expenseItem).toEqual(expected.expenseItem);
      expect(actual.quantity).toEqual(expected.quantity);
      expect(actual.freq).toEqual(expected.freq);
      expect(actual.unitCost).toEqual(expected.unitCost);
    });

    it('should correctly get board and lodging expenses of Bureau of Alternative Education', () => {
      const expected: Partial<ExpenseItem> = {
        expenseItem: 'Board and Lodging of Bureau of Alternative Education',
        quantity: 24,
        freq: 5,
        unitCost: 2000,
      };

      const result = budgetEstimate.getBoardAndLodging();
      const actual = result[3];

      expect(actual.expenseItem).toEqual(expected.expenseItem);
      expect(actual.quantity).toEqual(expected.quantity);
      expect(actual.freq).toEqual(expected.freq);
      expect(actual.unitCost).toEqual(expected.unitCost);
    });

    it('should correctly get board and lodging expenses of Other Offices', () => {
      const expected: Partial<ExpenseItem> = {
        expenseItem: 'Board and Lodging of Other Offices',
        quantity: 8,
        freq: 5,
        unitCost: 2000,
      };

      const result = budgetEstimate.getBoardAndLodging();
      const actual = result[4];

      expect(actual.expenseItem.trim()).toEqual(expected.expenseItem);
      expect(actual.quantity).toEqual(expected.quantity);
      expect(actual.freq).toEqual(expected.freq);
      expect(actual.unitCost).toEqual(expected.unitCost);
    });
  });

  describe('Travel Expenses', () => {
    describe('Travel Expenses of Participants', () => {
      it('should correctly get travel expenses of Participants from Region I', () => {
        const expected: Partial<ExpenseItem> = {
          expenseItem: 'Travel Expenses of Participants from Region I',
          quantity: 2,
          freq: 1,
          unitCost: 13900,
        };

        const info = budgetEstimate.getActivityInfo();

        const result = budgetEstimate.getTravelExpenses(info!.venue);
        const actual = result[0];

        expect(actual.expenseItem).toEqual(expected.expenseItem);
        expect(actual.quantity).toEqual(expected.quantity);
        expect(actual.freq).toEqual(expected.freq);
        expect(actual.unitCost).toEqual(expected.unitCost);
      });

      it('should correctly get travel expenses of Participants from Region II', () => {
        const expected: Partial<ExpenseItem> = {
          expenseItem: 'Travel Expenses of Participants from Region II',
          quantity: 2,
          freq: 1,
          unitCost: 14400,
        };

        const info = budgetEstimate.getActivityInfo();

        const result = budgetEstimate.getTravelExpenses(info!.venue);
        const actual = result[1];

        expect(actual.expenseItem).toEqual(expected.expenseItem);
        expect(actual.quantity).toEqual(expected.quantity);
        expect(actual.freq).toEqual(expected.freq);
        expect(actual.unitCost).toEqual(expected.unitCost);
      });

      it('should correctly get travel expenses of Participants from Batanes', () => {
        const expected: Partial<ExpenseItem> = {
          expenseItem: 'Travel Expenses of Participants from Batanes',
          quantity: 1,
          freq: 1,
          unitCost: 16400,
        };

        const info = budgetEstimate.getActivityInfo();

        const result = budgetEstimate.getTravelExpenses(info!.venue);
        const actual = result[2];

        expect(actual.expenseItem).toEqual(expected.expenseItem);
        expect(actual.quantity).toEqual(expected.quantity);
        expect(actual.freq).toEqual(expected.freq);
        expect(actual.unitCost).toEqual(expected.unitCost);
      });

      it('should correctly get travel expenses of Participants from Region III', () => {
        const expected: Partial<ExpenseItem> = {
          expenseItem: 'Travel Expenses of Participants from Region III',
          quantity: 4,
          freq: 1,
          unitCost: 13900,
        };

        const info = budgetEstimate.getActivityInfo();

        const result = budgetEstimate.getTravelExpenses(info!.venue);
        const actual = result[3];

        expect(actual.expenseItem).toEqual(expected.expenseItem);
        expect(actual.quantity).toEqual(expected.quantity);
        expect(actual.freq).toEqual(expected.freq);
        expect(actual.unitCost).toEqual(expected.unitCost);
      });

      it('should correctly get travel expenses of Participants from Calabarzon', () => {
        const expected: Partial<ExpenseItem> = {
          expenseItem: 'Travel Expenses of Participants from Calabarzon',
          quantity: 3,
          freq: 1,
          unitCost: 13900,
        };

        const info = budgetEstimate.getActivityInfo();

        const result = budgetEstimate.getTravelExpenses(info!.venue);
        const actual = result[4];

        expect(actual.expenseItem).toEqual(expected.expenseItem);
        expect(actual.quantity).toEqual(expected.quantity);
        expect(actual.freq).toEqual(expected.freq);
        expect(actual.unitCost).toEqual(expected.unitCost);
      });
      it('should correctly get travel expenses of Participants from Mimaropa', () => {
        const expected: Partial<ExpenseItem> = {
          expenseItem: 'Travel Expenses of Participants from Mimaropa',
          quantity: 1,
          freq: 1,
          unitCost: 13900,
        };

        const info = budgetEstimate.getActivityInfo();

        const result = budgetEstimate.getTravelExpenses(info!.venue);
        const actual = result[5];

        expect(actual.expenseItem).toEqual(expected.expenseItem);
        expect(actual.quantity).toEqual(expected.quantity);
        expect(actual.freq).toEqual(expected.freq);
        expect(actual.unitCost).toEqual(expected.unitCost);
      });

      it('should correctly get travel expenses of Participants from Palawan', () => {
        const expected: Partial<ExpenseItem> = {
          expenseItem: 'Travel Expenses of Participants from Palawan',
          quantity: 3,
          freq: 1,
          unitCost: 12100,
        };

        const info = budgetEstimate.getActivityInfo();

        const result = budgetEstimate.getTravelExpenses(info!.venue);
        const actual = result[6];

        expect(actual.expenseItem).toEqual(expected.expenseItem);
        expect(actual.quantity).toEqual(expected.quantity);
        expect(actual.freq).toEqual(expected.freq);
        expect(actual.unitCost).toEqual(expected.unitCost);
      });

      it('should correctly get travel expenses of Participants from Region V', () => {
        const expected: Partial<ExpenseItem> = {
          expenseItem: 'Travel Expenses of Participants from Region V',
          quantity: 2,
          freq: 1,
          unitCost: 13900,
        };

        const info = budgetEstimate.getActivityInfo();

        const result = budgetEstimate.getTravelExpenses(info!.venue);
        const actual = result[7];

        expect(actual.expenseItem).toEqual(expected.expenseItem);
        expect(actual.quantity).toEqual(expected.quantity);
        expect(actual.freq).toEqual(expected.freq);
        expect(actual.unitCost).toEqual(expected.unitCost);
      });

      it('should correctly get travel expenses of Participants from CAR', () => {
        const expected: Partial<ExpenseItem> = {
          expenseItem: 'Travel Expenses of Participants from CAR',
          quantity: 2,
          freq: 1,
          unitCost: 13900,
        };

        const info = budgetEstimate.getActivityInfo();

        const result = budgetEstimate.getTravelExpenses(info!.venue);
        const actual = result[8];

        expect(actual.expenseItem).toEqual(expected.expenseItem);
        expect(actual.quantity).toEqual(expected.quantity);
        expect(actual.freq).toEqual(expected.freq);
        expect(actual.unitCost).toEqual(expected.unitCost);
      });

      it('should correctly get travel expenses of Participants from NCR', () => {
        const expected: Partial<ExpenseItem> = {
          expenseItem: 'Travel Expenses of Participants from NCR',
          quantity: 3,
          freq: 1,
          unitCost: 13400,
        };

        const info = budgetEstimate.getActivityInfo();

        const result = budgetEstimate.getTravelExpenses(info!.venue);
        const actual = result[9];

        expect(actual.expenseItem).toEqual(expected.expenseItem);
        expect(actual.quantity).toEqual(expected.quantity);
        expect(actual.freq).toEqual(expected.freq);
        expect(actual.unitCost).toEqual(expected.unitCost);
      });

      it('should correctly get travel expenses of Participants from Region VI', () => {
        const expected: Partial<ExpenseItem> = {
          expenseItem: 'Travel Expenses of Participants from Region VI',
          quantity: 2,
          freq: 1,
          unitCost: 8400,
        };

        const info = budgetEstimate.getActivityInfo();

        const result = budgetEstimate.getTravelExpenses(info!.venue);
        const actual = result[10];

        expect(actual.expenseItem).toEqual(expected.expenseItem);
        expect(actual.quantity).toEqual(expected.quantity);
        expect(actual.freq).toEqual(expected.freq);
        expect(actual.unitCost).toEqual(expected.unitCost);
      });

      it('should correctly get travel expenses of Participants from Region VII', () => {
        const expected: Partial<ExpenseItem> = {
          expenseItem: 'Travel Expenses of Participants from Region VII',
          quantity: 3,
          freq: 1,
          unitCost: 8400,
        };

        const info = budgetEstimate.getActivityInfo();

        const result = budgetEstimate.getTravelExpenses(info!.venue);
        const actual = result[11];

        expect(actual.expenseItem).toEqual(expected.expenseItem);
        expect(actual.quantity).toEqual(expected.quantity);
        expect(actual.freq).toEqual(expected.freq);
        expect(actual.unitCost).toEqual(expected.unitCost);
      });

      it('should correctly get travel expenses of Participants from Region VIII', () => {
        const expected: Partial<ExpenseItem> = {
          expenseItem: 'Travel Expenses of Participants from Region VIII',
          quantity: 2,
          freq: 1,
          unitCost: 8400,
        };

        const info = budgetEstimate.getActivityInfo();

        const result = budgetEstimate.getTravelExpenses(info!.venue);
        const actual = result[12];

        expect(actual.expenseItem).toEqual(expected.expenseItem);
        expect(actual.quantity).toEqual(expected.quantity);
        expect(actual.freq).toEqual(expected.freq);
        expect(actual.unitCost).toEqual(expected.unitCost);
      });

      it('should correctly get travel expenses of Participants from Region IX', () => {
        const expected: Partial<ExpenseItem> = {
          expenseItem: 'Travel Expenses of Participants from Region IX',
          quantity: 1,
          freq: 1,
          unitCost: 8400,
        };

        const info = budgetEstimate.getActivityInfo();

        const result = budgetEstimate.getTravelExpenses(info!.venue);
        const actual = result[13];

        expect(actual.expenseItem).toEqual(expected.expenseItem);
        expect(actual.quantity).toEqual(expected.quantity);
        expect(actual.freq).toEqual(expected.freq);
        expect(actual.unitCost).toEqual(expected.unitCost);
      });

      it('should correctly get travel expenses of Participants from Region X', () => {
        const expected: Partial<ExpenseItem> = {
          expenseItem: 'Travel Expenses of Participants from Region X',
          quantity: 1,
          freq: 1,
          unitCost: 6400,
        };

        const info = budgetEstimate.getActivityInfo();

        const result = budgetEstimate.getTravelExpenses(info!.venue);
        const actual = result[14];

        expect(actual.expenseItem).toEqual(expected.expenseItem);
        expect(actual.quantity).toEqual(expected.quantity);
        expect(actual.freq).toEqual(expected.freq);
        expect(actual.unitCost).toEqual(expected.unitCost);
      });

      it('should correctly get travel expenses of Participants from Region XI', () => {
        const expected: Partial<ExpenseItem> = {
          expenseItem: 'Travel Expenses of Participants from Region XI',
          quantity: 2,
          freq: 1,
          unitCost: 6400,
        };

        const info = budgetEstimate.getActivityInfo();

        const result = budgetEstimate.getTravelExpenses(info!.venue);
        const actual = result[15];

        expect(actual.expenseItem).toEqual(expected.expenseItem);
        expect(actual.quantity).toEqual(expected.quantity);
        expect(actual.freq).toEqual(expected.freq);
        expect(actual.unitCost).toEqual(expected.unitCost);
      });

      it('should correctly get travel expenses of Participants from Region XII', () => {
        const expected: Partial<ExpenseItem> = {
          expenseItem: 'Travel Expenses of Participants from Region XII',
          quantity: 3,
          freq: 1,
          unitCost: 9400,
        };

        const info = budgetEstimate.getActivityInfo();

        const result = budgetEstimate.getTravelExpenses(info!.venue);
        const actual = result[16];

        expect(actual.expenseItem).toEqual(expected.expenseItem);
        expect(actual.quantity).toEqual(expected.quantity);
        expect(actual.freq).toEqual(expected.freq);
        expect(actual.unitCost).toEqual(expected.unitCost);
      });

      it('should correctly get travel expenses of Participants from CARAGA', () => {
        const expected: Partial<ExpenseItem> = {
          expenseItem: 'Travel Expenses of Participants from CARAGA',
          quantity: 1,
          freq: 1,
          unitCost: 2600,
        };

        const info = budgetEstimate.getActivityInfo();

        const result = budgetEstimate.getTravelExpenses(info!.venue);
        const actual = result[17];

        expect(actual.expenseItem).toEqual(expected.expenseItem);
        expect(actual.quantity).toEqual(expected.quantity);
        expect(actual.freq).toEqual(expected.freq);
        expect(actual.unitCost).toEqual(expected.unitCost);
      });
    });

    it('should correctly get travel expenses of resource persons', () => {
      const expected: Partial<ExpenseItem> = {
        expenseItem: 'Travel Expenses of Resource Persons',
        quantity: 8,
        freq: 1,
        unitCost: 13400,
      };

      const info = budgetEstimate.getActivityInfo();

      const result = budgetEstimate.getTravelExpenses(info!.venue);
      const actual = result[18];

      expect(actual.expenseItem).toEqual(expected.expenseItem);
      expect(actual.quantity).toEqual(expected.quantity);
      expect(actual.freq).toEqual(expected.freq);
      expect(actual.unitCost).toEqual(expected.unitCost);
    });

    it('should correctly get travel expenses of Technical Experts ( Editors/Validators/Illustrators)', () => {
      const expected: Partial<ExpenseItem> = {
        expenseItem:
          'Travel Expenses of Technical Experts ( Editors/Validators/Illustrators)',
        quantity: 4,
        freq: 1,
        unitCost: 13400,
      };

      const info = budgetEstimate.getActivityInfo();

      const result = budgetEstimate.getTravelExpenses(info!.venue);
      const actual = result[19];

      expect(actual.expenseItem).toEqual(expected.expenseItem);
      expect(actual.quantity).toEqual(expected.quantity);
      expect(actual.freq).toEqual(expected.freq);
      expect(actual.unitCost).toEqual(expected.unitCost);
    });

    it('should correctly get travel expenses of Bureau of Alternative Education', () => {
      const expected: Partial<ExpenseItem> = {
        expenseItem: 'Travel Expenses of Bureau of Alternative Education',
        quantity: 24,
        freq: 1,
        unitCost: 13400,
      };

      const info = budgetEstimate.getActivityInfo();

      const result = budgetEstimate.getTravelExpenses(info!.venue);
      const actual = result[20];

      expect(actual.expenseItem).toEqual(expected.expenseItem);
      expect(actual.quantity).toEqual(expected.quantity);
      expect(actual.freq).toEqual(expected.freq);
      expect(actual.unitCost).toEqual(expected.unitCost);
    });

    it('should correctly get travel expenses of Other Offices', () => {
      const expected: Partial<ExpenseItem> = {
        expenseItem: 'Travel Expenses of Other Offices',
        quantity: 8,
        freq: 1,
        unitCost: 13400,
      };

      const info = budgetEstimate.getActivityInfo();

      const result = budgetEstimate.getTravelExpenses(info!.venue);
      const actual = result[21];

      expect(actual.expenseItem.trim()).toEqual(expected.expenseItem);
      expect(actual.quantity).toEqual(expected.quantity);
      expect(actual.freq).toEqual(expected.freq);
      expect(actual.unitCost).toEqual(expected.unitCost);
    });
  });
  describe('Honorarium', () => {
    it('should correctly get honorarium of resource persons', () => {
      const expected: Partial<ExpenseItem> = {
        expenseItem: 'Honorarium of Resource Persons',
        quantity: 8,
        freq: 1,
        unitCost: 30000,
      };

      const result = budgetEstimate.getHonorarium();
      const actual = result[0];

      expect(actual.expenseItem).toEqual(expected.expenseItem);
      expect(actual.quantity).toEqual(expected.quantity);
      expect(actual.freq).toEqual(expected.freq);
      expect(actual.unitCost).toEqual(expected.unitCost);
    });

    it('should correctly get honorarium of Technical Experts ( Editors/Validators/Illustrators)', () => {
      const expected: Partial<ExpenseItem> = {
        expenseItem:
          'Honorarium of Technical Experts ( Editors/Validators/Illustrators)',
        quantity: 4,
        freq: 1,
        unitCost: 30000,
      };

      const result = budgetEstimate.getHonorarium();
      const actual = result[1];

      expect(actual.expenseItem).toEqual(expected.expenseItem);
      expect(actual.quantity).toEqual(expected.quantity);
      expect(actual.freq).toEqual(expected.freq);
      expect(actual.unitCost).toEqual(expected.unitCost);
    });
  });

  it('should correctly get supplies and materials', () => {
    const expected: Partial<ExpenseItem> = {
      expenseItem: 'Supplies and Materials',
      quantity: 85,
      freq: 1,
      unitCost: 100,
    };

    const result = budgetEstimate.getOtherExpenses();
    const actual = result[0];

    expect(actual.expenseItem.trim()).toEqual(expected.expenseItem);
    expect(actual.quantity).toEqual(expected.quantity);
    expect(actual.freq).toEqual(expected.freq);
    expect(actual.unitCost).toEqual(expected.unitCost);
  });

  it('should correctly get contingency', () => {
    const expected: Partial<ExpenseItem> = {
      expenseItem: 'Contingency',
      quantity: 1,
      freq: 1,
      unitCost: 4500,
    };

    const result = budgetEstimate.getOtherExpenses();
    const actual = result[1];

    expect(actual.expenseItem.trim()).toEqual(expected.expenseItem);
    expect(actual.quantity).toEqual(expected.quantity);
    expect(actual.freq).toEqual(expected.freq);
    expect(actual.unitCost).toEqual(expected.unitCost);
  });
  // Add more test cases for other methods as needed
});
