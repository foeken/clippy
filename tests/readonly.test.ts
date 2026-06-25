import { afterEach, expect, test } from 'bun:test';
import { assertReadWriteAllowed, isReadOnly } from '../src/lib/readonly.js';

const originalReadOnly = process.env.CLIPPY_READONLY;
const originalExit = process.exit;
const originalError = console.error;

afterEach(() => {
  if (originalReadOnly === undefined) {
    delete process.env.CLIPPY_READONLY;
  } else {
    process.env.CLIPPY_READONLY = originalReadOnly;
  }
  process.exit = originalExit;
  console.error = originalError;
});

test('isReadOnly accepts common truthy values', () => {
  for (const value of ['1', 'true', 'TRUE', 'yes', 'on']) {
    process.env.CLIPPY_READONLY = value;
    expect(isReadOnly()).toBe(true);
  }
});

test('isReadOnly rejects empty and false values', () => {
  for (const value of ['', '0', 'false', 'no', 'off']) {
    process.env.CLIPPY_READONLY = value;
    expect(isReadOnly()).toBe(false);
  }
});

test('assertReadWriteAllowed exits when read-only mode is enabled', () => {
  const errors: string[] = [];
  let exitCode: string | number | null | undefined;

  process.env.CLIPPY_READONLY = 'true';
  console.error = ((message?: unknown) => {
    errors.push(String(message));
  }) as typeof console.error;
  process.exit = ((code?: string | number | null | undefined) => {
    exitCode = code;
    throw new Error('process.exit');
  }) as typeof process.exit;

  expect(() => assertReadWriteAllowed('Sending email')).toThrow('process.exit');
  expect(exitCode).toBe(1);
  expect(errors).toEqual([
    'Read-only mode is enabled. Sending email is disabled.',
    'Unset CLIPPY_READONLY or remove --read-only to allow write operations.',
  ]);
});
