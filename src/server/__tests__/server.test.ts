import path from 'path';
import server from '../server';

// eslint-disable-next-line import/no-extraneous-dependencies
import request from 'supertest';

const cwd = process.cwd();
const endpoint = '/api/convert';
const formField = 'excelFile';
const dataDir = path.join(cwd, 'data');
const beFile = path.join(dataDir, 'be_test.xlsx');
const emFile = path.join(dataDir, 'em.xlsx');

describe('POST /api/convert', () => {
  it('response status should be 200 when official budget estimate was uploaded', async () => {
    const res = await request(server).post(endpoint).attach(formField, beFile);

    expect(res.status).toEqual(200);
  });

  it('response status should be 500 when invalid budget estimate was uploaded', async () => {
    const res = await request(server).post(endpoint).attach(formField, emFile);

    expect(res.status).toEqual(500);
  });
});
