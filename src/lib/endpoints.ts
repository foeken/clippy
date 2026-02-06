/**
 * Centralized Microsoft cloud endpoints.
 *
 * Set CLIPPY_CLOUD=gcc to use Office 365 US Government (GCC) endpoints.
 * Default is 'commercial' (regular Office 365).
 */

type CloudEnv = 'commercial' | 'gcc';

function getCloudEnv(): CloudEnv {
  const env = process.env.CLIPPY_CLOUD?.toLowerCase();
  if (env === 'gcc' || env === 'us') return 'gcc';
  return 'commercial';
}

const ENDPOINTS = {
  commercial: {
    outlookBase: 'https://outlook.office.com',
    loginBase: 'https://login.microsoftonline.com',
    graphBase: 'https://graph.microsoft.com',
  },
  gcc: {
    outlookBase: 'https://outlook.office365.us',
    loginBase: 'https://login.microsoftonline.com',
    graphBase: 'https://graph.microsoft.com',
  },
} as const;

const cloud = getCloudEnv();

export const OUTLOOK_BASE = ENDPOINTS[cloud].outlookBase;
export const LOGIN_BASE = ENDPOINTS[cloud].loginBase;
export const GRAPH_BASE = ENDPOINTS[cloud].graphBase;
export const OUTLOOK_API = `${OUTLOOK_BASE}/api/v2.0`;
export const OUTLOOK_SCOPE = `${OUTLOOK_BASE}/.default offline_access`;
export const CLOUD_ENV = cloud;
