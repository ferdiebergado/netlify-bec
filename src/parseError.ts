export class SheetParseError extends Error {
  public details: Record<string, any> | undefined;

  constructor(message: string, details?: Record<string, any>) {
    super(message);
    this.name = this.constructor.name; // Set the error name to the class name
    this.details = details;

    // Maintain proper stack trace (only on V8 engines)
    if (Error.captureStackTrace) {
      Error.captureStackTrace(this, this.constructor);
    }
  }
}
