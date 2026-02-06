import { Command } from 'commander';
import { resolveAuth, clearNeedsLogin } from '../lib/auth.js';
import { ensureConfigDir, saveConfig } from '../lib/config.js';
import { OUTLOOK_BASE } from '../lib/endpoints.js';

export const loginCommand = new Command('login')
  .description('Authenticate with OWA and validate session')
  .option('--token <token>', 'Use a specific token instead of extracting from browser')
  .option('-i, --interactive', 'Open browser to extract token automatically')
  .option('--no-headless', 'Show the browser window (don\'t run headless)')
  .option('--check', 'Only check if session is valid, do not save')
  .action(async (options: { token?: string; interactive?: boolean; headless?: boolean; check?: boolean }) => {
    console.log('Checking OWA session...');

    const result = await resolveAuth({
      token: options.token,
      interactive: options.interactive,
      headless: options.headless,
    });

    if (!result.success) {
      console.error(`\nError: ${result.error}`);
      console.error('\nTo authenticate:');
      console.error('1. Run with --interactive to open browser and extract token automatically');
      console.error('   bun run src/cli.ts login --interactive');
      console.error('\n2. Or set CLIPPY_TOKEN environment variable manually:');
      console.error(`   - Open ${OUTLOOK_BASE} in your browser`);
      console.error('   - Open DevTools (F12) → Network tab');
      console.error('   - Filter by "service.svc" and copy the Authorization header');
      console.error('   - export CLIPPY_TOKEN="eyJ..."');
      process.exit(1);
    }

    // Save config if not just checking
    if (!options.check) {
      try {
        await ensureConfigDir();
        await saveConfig({
          lastValidatedAt: new Date().toISOString(),
        });
        // Clear the needs-login marker so keepalive can resume
        await clearNeedsLogin();
      } catch (err) {
        // Non-fatal: continue even if config save fails
        console.warn('Warning: Could not save config');
      }
    }

    console.log('\n✓ Session valid');
    console.log('  Token: ***' + result.token!.slice(-8));
    console.log('\nYou are logged in to OWA. Run `clippy whoami` to see account details.');
  });
