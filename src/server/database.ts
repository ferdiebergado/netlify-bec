// eslint-disable-next-line import/no-extraneous-dependencies
import {
  MongoClient,
  Db,
  Collection,
  InsertOneResult,
  UpdateResult,
  ObjectId,
} from 'mongodb';
import { Document } from './types';
import config from './config';

const { uri, databaseName } = config.db;

async function insertDocument(
  document: Document,
  collectionName: string,
): Promise<ObjectId> {
  const client = new MongoClient(uri);
  console.log('Connecting to the database...');

  try {
    // Connect to MongoDB
    await client.connect();
    console.log('Connected to the database.');

    // Access the database
    const database: Db = client.db(databaseName);

    // Access the collection
    const collection: Collection<Document> =
      database.collection(collectionName);

    // Insert the document into the collection
    const result: InsertOneResult<Document> =
      await collection.insertOne(document);

    console.log('Document saved.');

    // Check the result
    return result.insertedId;
  } finally {
    // Close the connection
    await client.close();
    console.log('Connection closed');
  }
}

async function updateDocument(
  collectionName: string,
  filter: Record<string, any>,
  update: Record<string, any>,
): Promise<UpdateResult> {
  const client = new MongoClient(uri);

  try {
    await client.connect();
    console.log('Connected to the database.');

    const database = client.db(databaseName);
    const collection = database.collection<Document>(collectionName);
    const result = await collection.updateOne(filter, update);
    console.log('Document updated.');

    return result;
  } finally {
    await client.close();
    console.log('Connection closed');
  }
}

export { insertDocument, updateDocument };
