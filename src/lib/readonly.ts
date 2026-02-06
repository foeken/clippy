export function isReadOnly(): boolean {
  const value = process.env.CLIPPY_READONLY?.toLowerCase();
  return value === '1' || value === 'true' || value === 'yes' || value === 'on';
}

export function assertReadWriteAllowed(action: string): void {
  if (!isReadOnly()) return;

  console.error(`Read-only mode is enabled. ${action} is disabled.`);
  console.error('Unset CLIPPY_READONLY or remove --read-only to allow write operations.');
  process.exit(1);
}
