import { chromium } from 'playwright';
import { setTimeout as sleep } from 'timers/promises';
import { homedir } from 'os';
import { join } from 'path';
import { readFile, writeFile, mkdir, unlink, open, readlink, rm } from 'fs/promises';
import { validateSession } from './owa-client.js';
import { OUTLOOK_BASE, LOGIN_BASE, GRAPH_BASE, OUTLOOK_SCOPE } from './endpoints.js';

/**
 * Clean up stale Chrome SingletonLock if the owning process is no longer running.
 * Chrome creates a symlink at `<profile>/SingletonLock` pointing to `hostname-pid`.
 * If that process is dead, we can safely remove the lock to allow a new instance.
 */
async function cleanStaleChromeProfile(profileDir: string): Promise<void> {
  const singletonLock = join(profileDir, 'SingletonLock');

  try {
    // SingletonLock is a symlink pointing to "hostname-pid"
    const target = await readlink(singletonLock);
    const match = target.match(/-(\d+)$/);

    if (match) {
      const pid = parseInt(match[1], 10);

      // Check if process is still running
      const isRunning = isProcessRunning(pid);

      if (!isRunning) {
        // Process is dead, clean up stale lock files
        console.log(`Cleaning up stale browser lock (dead PID ${pid})...`);

        // Remove all singleton-related files
        const filesToRemove = ['SingletonLock', 'SingletonCookie', 'SingletonSocket'];
        for (const file of filesToRemove) {
          try {
            await rm(join(profileDir, file), { force: true });
          } catch {
            // Ignore errors for individual files
          }
        }
      }
    }
  } catch {
    // SingletonLock doesn't exist or isn't a symlink - nothing to clean
  }
}

/**
 * Check if a process with the given PID is currently running.
 */
function isProcessRunning(pid: number): boolean {
  try {
    // Sending signal 0 checks if process exists without actually sending a signal
    process.kill(pid, 0);
    return true;
  } catch (err: unknown) {
    // ESRCH = No such process (not running)
    // EPERM = Process exists but we don't have permission (still running)
    if (err && typeof err === 'object' && 'code' in err) {
      return err.code === 'EPERM';
    }
    return false;
  }
}

const LOCK_FILE = join(homedir(), '.config', 'clippy', 'browser.lock');
const LOCK_TIMEOUT = 60000; // Consider lock stale after 60s
const NEEDS_LOGIN_FILE = join(homedir(), '.config', 'clippy', 'needs-login');

/**
 * Check if a needs-login marker exists (session expired, needs manual re-auth).
 */
export async function needsLogin(): Promise<boolean> {
  try {
    await readFile(NEEDS_LOGIN_FILE);
    return true;
  } catch {
    return false;
  }
}

/**
 * Set the needs-login marker (called when session expires).
 */
export async function setNeedsLogin(reason: string): Promise<void> {
  try {
    const cacheDir = join(homedir(), '.config', 'clippy');
    await mkdir(cacheDir, { recursive: true });
    await writeFile(NEEDS_LOGIN_FILE, `${new Date().toISOString()}: ${reason}`, 'utf-8');
  } catch {
    // Ignore write errors
  }
}

/**
 * Clear the needs-login marker (called after successful login).
 */
export async function clearNeedsLogin(): Promise<void> {
  try {
    await unlink(NEEDS_LOGIN_FILE);
  } catch {
    // Ignore if doesn't exist
  }
}

interface LockHandle {
  release: () => Promise<void>;
}

async function acquireLock(timeout: number = 10000): Promise<LockHandle | null> {
  const startTime = Date.now();

  while (Date.now() - startTime < timeout) {
    try {
      // Check if lock exists and is stale
      try {
        const lockData = await readFile(LOCK_FILE, 'utf-8');
        const lockTime = parseInt(lockData, 10);
        if (Date.now() - lockTime > LOCK_TIMEOUT) {
          // Stale lock, remove it
          await unlink(LOCK_FILE);
        } else {
          // Lock is held, wait and retry
          await new Promise(r => setTimeout(r, 500));
          continue;
        }
      } catch {
        // Lock doesn't exist, proceed
      }

      // Try to create lock file (atomic via exclusive create)
      await mkdir(join(homedir(), '.config', 'clippy'), { recursive: true });
      const fd = await open(LOCK_FILE, 'wx');
      await fd.write(Date.now().toString());
      await fd.close();

      return {
        release: async () => {
          try {
            await unlink(LOCK_FILE);
          } catch {
            // Ignore
          }
        },
      };
    } catch (err: unknown) {
      if (err && typeof err === 'object' && 'code' in err && err.code === 'EEXIST') {
        // Lock file exists, wait and retry
        await new Promise(r => setTimeout(r, 500));
        continue;
      }
      // Other error, try again
      await new Promise(r => setTimeout(r, 100));
    }
  }

  return null; // Couldn't acquire lock
}

export interface AuthResult {
  success: boolean;
  token?: string;
  graphToken?: string;
  refreshToken?: string;
  error?: string;
}

export interface PlaywrightTokenResult {
  success: boolean;
  token?: string;
  graphToken?: string;
  refreshToken?: string;
  error?: string;
}

interface CachedToken {
  token: string;
  graphToken?: string;
  refreshToken?: string;
  expiresAt: number;
}

const TOKEN_CACHE_FILE = join(homedir(), '.config', 'clippy', 'token-cache.json');
const REFRESH_THRESHOLD = 5 * 60 * 1000; // Refresh if less than 5 minutes remaining

function getJwtExpiration(token: string): number | null {
  try {
    const parts = token.split('.');
    if (parts.length !== 3) return null;
    const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString());
    return payload.exp ? payload.exp * 1000 : null;
  } catch {
    return null;
  }
}

async function getCachedToken(): Promise<{ cached: CachedToken; needsRefresh: boolean } | null> {
  try {
    const data = await readFile(TOKEN_CACHE_FILE, 'utf-8');
    const cached = JSON.parse(data) as CachedToken;
    const now = Date.now();

    // If token is expired but we have a refresh token, try to refresh
    if (cached.expiresAt <= now) {
      if (cached.refreshToken) {
        console.log('[getCachedToken] Access token expired, attempting refresh...');
        const refreshResult = await refreshAccessToken(cached.refreshToken);
        
        if (refreshResult.success && refreshResult.accessToken) {
          // Save the new tokens
          const newRefreshToken = refreshResult.refreshToken || cached.refreshToken;
          await setCachedToken(refreshResult.accessToken, cached.graphToken, newRefreshToken);
          
          // Return the refreshed token
          const newExpiry = getJwtExpiration(refreshResult.accessToken) || (Date.now() + 55 * 60 * 1000);
          return {
            cached: {
              token: refreshResult.accessToken,
              graphToken: cached.graphToken,
              refreshToken: newRefreshToken,
              expiresAt: newExpiry,
            },
            needsRefresh: false,
          };
        } else {
          console.error('[getCachedToken] Refresh failed:', refreshResult.error);
        }
      }
      return null; // Expired and couldn't refresh
    }

    // Check if we need to refresh soon
    const needsRefresh = (cached.expiresAt - now) < REFRESH_THRESHOLD;
    return { cached, needsRefresh };
  } catch {
    // No cache or invalid cache
  }
  return null;
}

// Microsoft's first-party Outlook Web App client ID
const OUTLOOK_CLIENT_ID = '9199bf20-a13f-4107-85dc-02114787ef48';

/**
 * Use a refresh token to get a new access token without browser.
 * This uses Microsoft's first-party Outlook client ID.
 */
export async function refreshAccessToken(refreshToken: string): Promise<{
  success: boolean;
  accessToken?: string;
  refreshToken?: string;
  expiresIn?: number;
  error?: string;
}> {
  try {
    console.log('[refreshAccessToken] Attempting token refresh via OAuth...');
    
    const response = await fetch(`${LOGIN_BASE}/common/oauth2/v2.0/token`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Origin': OUTLOOK_BASE, // Required for SPA tokens
      },
      body: new URLSearchParams({
        client_id: OUTLOOK_CLIENT_ID,
        grant_type: 'refresh_token',
        refresh_token: refreshToken,
        scope: OUTLOOK_SCOPE,
      }).toString(),
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error('[refreshAccessToken] Failed:', response.status, errorText);
      return {
        success: false,
        error: `Token refresh failed: ${response.status} - ${errorText}`,
      };
    }

    const json = await response.json() as {
      access_token?: string;
      refresh_token?: string;
      expires_in?: number;
    };

    if (json.access_token) {
      console.log('[refreshAccessToken] ✅ Got new access token via refresh!');
      return {
        success: true,
        accessToken: json.access_token,
        refreshToken: json.refresh_token, // Microsoft may rotate refresh tokens
        expiresIn: json.expires_in,
      };
    }

    return {
      success: false,
      error: 'No access token in response',
    };
  } catch (err) {
    console.error('[refreshAccessToken] Exception:', err);
    return {
      success: false,
      error: err instanceof Error ? err.message : 'Unknown error',
    };
  }
}

export async function setCachedToken(token: string, graphToken?: string, refreshToken?: string): Promise<void> {
  try {
    const cacheDir = join(homedir(), '.config', 'clippy');
    await mkdir(cacheDir, { recursive: true });

    // Use actual JWT expiration time
    const expiresAt = getJwtExpiration(token) || (Date.now() + 55 * 60 * 1000);

    // Preserve existing refresh token if not provided
    let existingRefreshToken: string | undefined;
    if (!refreshToken) {
      try {
        const existing = await readFile(TOKEN_CACHE_FILE, 'utf-8');
        const parsed = JSON.parse(existing) as CachedToken;
        existingRefreshToken = parsed.refreshToken;
      } catch {
        // No existing cache
      }
    }

    const cached: CachedToken = {
      token,
      graphToken,
      refreshToken: refreshToken || existingRefreshToken,
      expiresAt,
    };
    await writeFile(TOKEN_CACHE_FILE, JSON.stringify(cached, null, 2), 'utf-8');
    
    if (refreshToken) {
      console.log('[setCachedToken] Refresh token saved!');
    }
  } catch {
    // Ignore cache write errors
  }
}

/**
 * Extract Bearer token by launching a browser and intercepting OWA requests.
 * Uses a persistent profile so the user only needs to log in once.
 * Tries headless first, then falls back to visible browser if login is needed.
 */
export async function extractTokenViaPlaywright(
  options: { headless?: boolean; timeout?: number; userDataDir?: string; fallbackToVisible?: boolean } = {}
): Promise<PlaywrightTokenResult> {
  const { headless = true, timeout = 15000, userDataDir, fallbackToVisible = true } = options;

  console.log(`[extractToken] Starting with headless=${headless}, timeout=${timeout}`);

  // Try headless first (fast path for already logged-in users)
  const result = await tryExtractToken(headless, timeout, userDataDir);

  console.log(`[extractToken] Headless result: success=${result.success}, error=${result.error || 'none'}`);

  if (result.success) {
    return result;
  }

  // If headless failed and we haven't tried visible yet, retry with visible browser
  if (headless && fallbackToVisible) {
    console.log('Session not found. Opening browser for login...');
    console.log('[extractToken] Falling back to visible browser with 120s timeout...');
    const visibleResult = await tryExtractToken(false, 120000, userDataDir);
    console.log(`[extractToken] Visible result: success=${visibleResult.success}, error=${visibleResult.error || 'none'}`);
    return visibleResult;
  }

  return result;
}

// Path to storage state file (cookies + localStorage as JSON)
function getStorageStatePath(): string {
  return join(homedir(), '.config', 'clippy', 'storage-state.json');
}

async function tryExtractToken(
  headless: boolean,
  timeout: number,
  userDataDirOverride?: string
): Promise<PlaywrightTokenResult> {
  console.log(`[tryExtract] Starting: headless=${headless}, timeout=${timeout}`);
  
  // Acquire lock to prevent concurrent browser access
  console.log(`[tryExtract] Acquiring lock (${headless ? 5000 : 30000}ms timeout)...`);
  const lock = await acquireLock(headless ? 5000 : 30000);
  if (!lock) {
    console.log('[tryExtract] Failed to acquire lock');
    return {
      success: false,
      error: 'Another browser session is in progress, please wait',
    };
  }
  console.log('[tryExtract] Lock acquired');

  let browser;
  let context;
  try {
    // Check if we have saved storage state (cookies as JSON)
    const storageStatePath = getStorageStatePath();
    let hasStorageState = false;
    try {
      await readFile(storageStatePath);
      hasStorageState = true;
      console.log('[tryExtract] Found storage state file');
    } catch {
      console.log('[tryExtract] No storage state file');
    }

    console.log(`[tryExtract] Launching browser (headless=${headless})...`);
    // Launch regular browser (not persistent context - storageState works better this way)
    browser = await chromium.launch({
      headless,
      channel: 'chrome',
      args: [
        '--disable-blink-features=AutomationControlled',
      ],
    });
    console.log('[tryExtract] Browser launched');

    // Create context with storage state if available
    console.log('[tryExtract] Creating context...');
    context = await browser.newContext(
      hasStorageState ? { storageState: storageStatePath } : undefined
    );
    console.log('[tryExtract] Context created');

    const page = await context.newPage();
    console.log('[tryExtract] Page created');

    // Debug: track page/browser close events
    page.on('close', () => console.log('[tryExtract] EVENT: page closed'));
    context.on('close', () => console.log('[tryExtract] EVENT: context closed'));
    browser.on('disconnected', () => console.log('[tryExtract] EVENT: browser disconnected'));

    let capturedToken: string | null = null;
    let capturedGraphToken: string | null = null;
    let capturedRefreshToken: string | null = null;

    // Intercept OAuth token endpoint responses to capture refresh tokens
    page.on('response', async response => {
      const url = response.url();
      
      // Log all Microsoft auth-related responses for debugging
      if (url.includes('microsoftonline.com') || url.includes('login.microsoft')) {
        console.log(`[tryExtract] Auth response: ${response.status()} ${url.substring(0, 100)}...`);
      }
      
      // Capture refresh token from OAuth token endpoint
      if (url.includes(LOGIN_BASE) && url.includes('/oauth2/') && url.includes('/token')) {
        console.log('[tryExtract] 🔑 Token endpoint response detected!');
        try {
          const text = await response.text();
          console.log('[tryExtract] Token response length:', text.length);
          const json = JSON.parse(text);
          if (json.refresh_token && !capturedRefreshToken) {
            capturedRefreshToken = json.refresh_token;
            console.log('[tryExtract] 🎉 Captured refresh token from OAuth response!');
            
            // Save immediately so we don't lose it if something hangs later
            if (json.access_token) {
              console.log('[tryExtract] Saving tokens immediately...');
              setCachedToken(json.access_token, undefined, json.refresh_token)
                .then(() => console.log('[tryExtract] Tokens saved to disk!'))
                .catch(err => console.log('[tryExtract] Token save error:', err));
            }
          }
          if (json.access_token && !capturedToken) {
            capturedToken = json.access_token;
            console.log('[tryExtract] Captured access token from OAuth response');
          }
        } catch (err) {
          console.log('[tryExtract] Token response parse error:', err);
        }
      }
    });

    // Intercept requests to capture Bearer tokens
    page.on('request', request => {
      const url = request.url();
      const headers = request.headers();
      const authHeader = headers['authorization'];

      if (authHeader && authHeader.startsWith('Bearer ')) {
        const token = authHeader.replace('Bearer ', '');

        // Capture Outlook token
        if (url.includes(OUTLOOK_BASE) && !capturedToken) {
          capturedToken = token;
        }

        // Capture Graph token
        if (url.includes(GRAPH_BASE) && !capturedGraphToken) {
          capturedGraphToken = token;
        }
      }
    });

    if (!headless) {
      console.log('Please complete the login process in the browser...');
    }

    console.log(`[tryExtract] Navigating to Outlook (timeout=${timeout}ms)...`);
    await page.goto(`${OUTLOOK_BASE}/mail/`, {
      waitUntil: 'domcontentloaded',
      timeout,
    });
    console.log(`[tryExtract] Page loaded (domcontentloaded). URL: ${page.url()}`);

    // Wait for token to be captured (max timeout)
    console.log(`[tryExtract] Waiting for token capture (max ${timeout}ms)...`);
    const startTime = Date.now();
    while (!capturedToken && (Date.now() - startTime) < timeout) {
      await page.waitForTimeout(500);
      // Log progress every 10 seconds
      const elapsed = Date.now() - startTime;
      if (elapsed % 10000 < 500) {
        console.log(`[tryExtract] Still waiting for token... ${Math.round(elapsed / 1000)}s elapsed`);
      }
    }
    console.log(`[tryExtract] Token wait finished. capturedToken=${!!capturedToken}`);

    // If we have the Outlook token, wait a bit more to try to capture Graph token
    if (capturedToken && !capturedGraphToken) {
      // Try to trigger Graph API calls by interacting with the page
      const graphWaitTime = headless ? 3000 : 5000;
      const graphStart = Date.now();
      while (!capturedGraphToken && (Date.now() - graphStart) < graphWaitTime) {
        await page.waitForTimeout(500);
      }
    }

    // Save storage state (cookies as JSON) to persist session across restarts
    // This bypasses Chrome's cookie encryption which doesn't work with Playwright's mock keychain
    if (capturedToken) {
      console.log('[tryExtract] Saving storage state...');
      try {
        await Promise.race([
          context.storageState({ path: getStorageStatePath() }),
          new Promise((_, reject) => setTimeout(() => reject(new Error('Storage state timeout')), 10000))
        ]);
        console.log('[tryExtract] Storage state saved');
      } catch (err) {
        console.log('[tryExtract] Storage state save failed:', err);
        // Non-fatal: continue even if storage state save fails
      }
    }

    console.log('[tryExtract] Closing browser...');
    try {
      await Promise.race([
        browser.close(),
        new Promise((_, reject) => setTimeout(() => reject(new Error('Browser close timeout')), 10000))
      ]);
    } catch (err) {
      console.log('[tryExtract] Browser close issue:', err);
    }
    await lock.release();
    console.log('[tryExtract] Done, returning result');

    if (capturedToken) {
      return { success: true, token: capturedToken, graphToken: capturedGraphToken || undefined, refreshToken: capturedRefreshToken || undefined };
    }

    return {
      success: false,
      error: headless
        ? 'No active session found'
        : 'Timeout: No Bearer token captured. Make sure you completed the login.'
    };
  } catch (err) {
    console.error('[tryExtract] EXCEPTION:', err instanceof Error ? err.message : err);
    console.error('[tryExtract] Stack:', err instanceof Error ? err.stack : 'no stack');
    if (browser) {
      await browser.close();
    }
    await lock.release();
    return {
      success: false,
      error: err instanceof Error ? err.message : 'Unknown error during token extraction',
    };
  }
}

// Default profile directory - can be overridden via CLIPPY_PROFILE_DIR env var
export function getProfileDir(): string {
  return process.env.CLIPPY_PROFILE_DIR || join(homedir(), '.config', 'clippy', 'browser-profile');
}

export async function startKeepaliveSession(options: { intervalMinutes: number; headless?: boolean }): Promise<void> {
  const { intervalMinutes, headless = false } = options;

  // Check if needs-login marker exists - if so, exit immediately without opening browser
  if (await needsLogin()) {
    console.log(`[${new Date().toISOString()}] Session expired. Run 'clippy login --interactive' to re-authenticate.`);
    console.log(`[${new Date().toISOString()}] Keepalive will resume automatically after login.`);
    // Exit with code 0 so launchd doesn't spam restarts (ThrottleInterval still applies)
    // The marker file prevents us from doing anything until manual login clears it
    process.exit(0);
  }

  // Check if we have saved storage state (cookies as JSON)
  const storageStatePath = getStorageStatePath();
  let hasStorageState = false;
  try {
    await readFile(storageStatePath);
    hasStorageState = true;
  } catch {
    // No storage state file yet
  }

  // Launch regular browser (not persistent context - storageState works better this way)
  const browser = await chromium.launch({
    headless,
    channel: 'chrome',
    args: [
      '--disable-blink-features=AutomationControlled',
    ],
  });

  // Create context with storage state if available
  const context = await browser.newContext(
    hasStorageState ? { storageState: storageStatePath } : undefined
  );

  const page = await context.newPage();

  let lastToken: string | null = null;
  let lastGraphToken: string | null = null;
  let consecutiveFailures = 0;
  const MAX_CONSECUTIVE_FAILURES = 3;

  page.on('request', request => {
    const url = request.url();
    const headers = request.headers();
    const authHeader = headers['authorization'];

    if (authHeader && authHeader.startsWith('Bearer ')) {
      const token = authHeader.replace('Bearer ', '');

      if (url.includes(OUTLOOK_BASE)) {
        lastToken = token;
      }
      if (url.includes(GRAPH_BASE)) {
        lastGraphToken = token;
      }
    }
  });

  // Helper: check if we landed on a login page
  const isLoginPage = (url: string): boolean => {
    return url.includes('/login') || 
           url.includes(LOGIN_BASE) ||
           url.includes('login.live.com') ||
           url.includes('/oauth2/') ||
           url.includes('/common/oauth2/');
  };

  // Helper: simulate user activity by clicking on UI elements
  const simulateActivity = async (): Promise<void> => {
    try {
      // Try clicking on Inbox folder to simulate real user interaction
      const inboxSelectors = [
        '[data-folder-name="inbox"]',
        '[aria-label*="Inbox"]',
        '[aria-label*="Postvak IN"]', // Dutch
        'button[title*="Inbox"]',
        'div[role="treeitem"][aria-label*="Inbox"]',
      ];
      
      for (const selector of inboxSelectors) {
        const element = await page.$(selector);
        if (element) {
          await element.click().catch(() => {});
          console.log(`[${new Date().toISOString()}] Clicked inbox element: ${selector}`);
          await sleep(1000);
          return;
        }
      }
      
      // Fallback: try clicking the first email in the list to trigger activity
      const emailSelectors = [
        '[role="option"]',
        '[aria-label*="message"]',
        '.customScrollBar div[role="listbox"] > div:first-child',
      ];
      
      for (const selector of emailSelectors) {
        const element = await page.$(selector);
        if (element) {
          await element.click().catch(() => {});
          console.log(`[${new Date().toISOString()}] Clicked email element: ${selector}`);
          await sleep(1000);
          return;
        }
      }
      
      console.log(`[${new Date().toISOString()}] No clickable elements found, will rely on page reload`);
    } catch (err) {
      console.log(`[${new Date().toISOString()}] Activity simulation failed: ${err instanceof Error ? err.message : 'unknown'}`);
    }
  };

  console.log(`Opening Outlook session (headless=${headless ? 'true' : 'false'})...`);
  await page.goto(`${OUTLOOK_BASE}/mail/`, { waitUntil: 'domcontentloaded', timeout: 60000 });

  // Immediately check if we landed on a login page
  const initialUrl = page.url();
  if (isLoginPage(initialUrl)) {
    console.error(`[${new Date().toISOString()}] CRITICAL: Landed on login page! Session expired.`);
    console.error(`[${new Date().toISOString()}] Run 'clippy login --interactive' to re-authenticate.`);
    await setNeedsLogin('Landed on login page at startup');
    await browser.close();
    process.exit(0);
  }

  console.log(`Keepalive active. Refreshing every ${intervalMinutes} minutes.`);

  const healthFile = join(homedir(), '.config', 'clippy', 'keepalive-health.txt');

  // Main keepalive loop
  while (true) {
    // Give some time for requests to fire and tokens to be captured
    await sleep(2000);

    // Check current URL for login redirect
    const currentUrl = page.url();
    if (isLoginPage(currentUrl)) {
      console.error(`[${new Date().toISOString()}] CRITICAL: Redirected to login page! Session expired.`);
      console.error(`[${new Date().toISOString()}] Run 'clippy login --interactive' to re-authenticate.`);
      await setNeedsLogin('Redirected to login page during keepalive');
      await browser.close();
      process.exit(0);
    }

    if (lastToken) {
      // Validate the token before caching it
      const isValid = await validateSession(lastToken);
      
      if (isValid) {
        await setCachedToken(lastToken, lastGraphToken || undefined);
        // Write health file for external monitoring
        await writeFile(healthFile, new Date().toISOString(), 'utf-8').catch(() => {});
        // Save storage state (cookies as JSON) to persist session
        await context.storageState({ path: getStorageStatePath() }).catch(() => {});
        consecutiveFailures = 0; // Reset on success
        console.log(`[${new Date().toISOString()}] Token validated and cached successfully`);
      } else {
        consecutiveFailures++;
        console.error(`[${new Date().toISOString()}] Token validation failed (attempt ${consecutiveFailures}/${MAX_CONSECUTIVE_FAILURES})`);
        
        if (consecutiveFailures >= MAX_CONSECUTIVE_FAILURES) {
          console.error(`[${new Date().toISOString()}] Too many consecutive failures, exiting for launchd restart...`);
          await browser.close();
          process.exit(1);
        }
      }
    } else {
      consecutiveFailures++;
      console.error(`[${new Date().toISOString()}] No token captured (attempt ${consecutiveFailures}/${MAX_CONSECUTIVE_FAILURES})`);
      
      if (consecutiveFailures >= MAX_CONSECUTIVE_FAILURES) {
        console.error(`[${new Date().toISOString()}] Too many consecutive failures, exiting for launchd restart...`);
        await browser.close();
        process.exit(1);
      }
    }

    await sleep(intervalMinutes * 60 * 1000);

    // Reset tokens before reload to capture fresh ones
    lastToken = null;
    lastGraphToken = null;

    try {
      // First reload the page
      await page.reload({ waitUntil: 'domcontentloaded', timeout: 60000 });
      
      // Check if reload landed us on login page
      const afterReloadUrl = page.url();
      if (isLoginPage(afterReloadUrl)) {
        console.error(`[${new Date().toISOString()}] CRITICAL: Reload landed on login page! Session expired.`);
        console.error(`[${new Date().toISOString()}] Run 'clippy login --interactive' to re-authenticate.`);
        await setNeedsLogin('Reload landed on login page');
        await browser.close();
        process.exit(0);
      }
      
      // Simulate user activity to keep session warm
      await sleep(2000); // Wait for page to settle
      await simulateActivity();
      
    } catch {
      // If reload fails, try to navigate again
      try {
        await page.goto(`${OUTLOOK_BASE}/mail/`, { waitUntil: 'domcontentloaded', timeout: 60000 });
        
        // Check if navigation landed us on login page
        const afterNavUrl = page.url();
        if (isLoginPage(afterNavUrl)) {
          console.error(`[${new Date().toISOString()}] CRITICAL: Navigation landed on login page! Session expired.`);
          console.error(`[${new Date().toISOString()}] Run 'clippy login --interactive' to re-authenticate.`);
          await setNeedsLogin('Navigation landed on login page');
          await browser.close();
          process.exit(0);
        }
        
        await sleep(2000);
        await simulateActivity();
      } catch {
        // ignore, will retry on next loop
      }
    }
  }
}

export async function resolveAuth(options: {
  token?: string;
  interactive?: boolean;
  headless?: boolean;
}): Promise<AuthResult> {
  const { token: cliToken, interactive = false, headless } = options;

  // Priority 1: CLI flag
  if (cliToken) {
    const isValid = await validateSession(cliToken);
    if (isValid) {
      return { success: true, token: cliToken };
    }
    return { success: false, error: 'Provided token is invalid or expired' };
  }

  // Priority 2: Environment variable
  const envToken = process.env.CLIPPY_TOKEN;
  if (envToken) {
    const isValid = await validateSession(envToken);
    if (isValid) {
      return { success: true, token: envToken };
    }
    return {
      success: false,
      error: 'CLIPPY_TOKEN environment variable contains invalid or expired token',
    };
  }

  // Priority 3: Cached token (fast path - no browser needed)
  const cachedResult = await getCachedToken();
  if (cachedResult) {
    const { cached, needsRefresh } = cachedResult;
    const isValid = await validateSession(cached.token);

    if (isValid) {
      // If token is expiring soon, refresh in background (silent)
      if (needsRefresh && interactive) {
        extractTokenViaPlaywright({ headless: true, timeout: 10000 })
          .then(result => {
            if (result.success && result.token) {
              setCachedToken(result.token, result.graphToken);
            }
          })
          .catch(() => {}); // Ignore refresh errors
      }

      return {
        success: true,
        token: cached.token,
        graphToken: cached.graphToken,
      };
    }
    // Cache is stale, will re-extract below
  }

  // Priority 4: Interactive Playwright extraction
  if (interactive) {
    const playwrightResult = await extractTokenViaPlaywright({
      headless: headless !== undefined ? headless : true,
      fallbackToVisible: headless !== false,
    });
    if (playwrightResult.success && playwrightResult.token) {
      const isValid = await validateSession(playwrightResult.token);
      if (isValid) {
        // Cache the token for future use (including refresh token if captured)
        await setCachedToken(playwrightResult.token, playwrightResult.graphToken, playwrightResult.refreshToken);
        return {
          success: true,
          token: playwrightResult.token,
          graphToken: playwrightResult.graphToken,
          refreshToken: playwrightResult.refreshToken,
        };
      }
      return {
        success: false,
        error: 'Extracted token is invalid or expired',
      };
    }
    return {
      success: false,
      error: playwrightResult.error || 'Failed to extract token via browser',
    };
  }

  // No token available and not interactive
  return {
    success: false,
    error: 'No token available. Set CLIPPY_TOKEN env var or run with --interactive to extract via browser.',
  };
}
