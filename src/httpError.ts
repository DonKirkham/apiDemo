// An Error that carries an HTTP status code. Throw it anywhere in the request
// path; the handler wrapper in functions/helpers.ts turns it into a response
// with that status. Any other error becomes a 500.
export class HttpError extends Error {
  constructor(
    public readonly status: number,
    message: string
  ) {
    super(message);
    this.name = 'HttpError';
  }
}
