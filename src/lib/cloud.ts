export type CloudEnvironment = 'commercial' | 'gcc';

export interface CloudConfig {
  environment: CloudEnvironment;
  label: string;
  authorityHost: string;
  ewsResource: string;
  ewsEndpoint: string;
}

const CLOUD_CONFIGS: Record<CloudEnvironment, CloudConfig> = {
  commercial: {
    environment: 'commercial',
    label: 'Microsoft 365 commercial',
    authorityHost: 'https://login.microsoftonline.com',
    ewsResource: 'https://outlook.office365.com',
    ewsEndpoint: 'https://outlook.office365.com/EWS/Exchange.asmx',
  },
  gcc: {
    environment: 'gcc',
    label: 'Microsoft 365 US Government',
    authorityHost: 'https://login.microsoftonline.us',
    ewsResource: 'https://outlook.office365.us',
    ewsEndpoint: 'https://outlook.office365.us/EWS/Exchange.asmx',
  },
};

function trimTrailingSlash(value: string): string {
  return value.replace(/\/+$/, '');
}

function envOverride(name: string): string | undefined {
  const value = process.env[name]?.trim();
  return value ? trimTrailingSlash(value) : undefined;
}

export function resolveCloudEnvironment(value = process.env.CLIPPY_CLOUD): CloudEnvironment {
  const normalized = value?.trim().toLowerCase();

  if (!normalized || ['commercial', 'global', 'public', 'worldwide'].includes(normalized)) {
    return 'commercial';
  }

  if (['gcc', 'gcc-high', 'gcch', 'usgov', 'us-gov', 'government'].includes(normalized)) {
    return 'gcc';
  }

  throw new Error(`Unsupported CLIPPY_CLOUD "${value}". Use "commercial" or "gcc".`);
}

export function getCloudConfig(): CloudConfig {
  const base = CLOUD_CONFIGS[resolveCloudEnvironment()];

  return {
    ...base,
    authorityHost: envOverride('EWS_AUTHORITY_HOST') || base.authorityHost,
    ewsResource: envOverride('EWS_RESOURCE') || base.ewsResource,
    ewsEndpoint: process.env.EWS_ENDPOINT?.trim() || base.ewsEndpoint,
  };
}

export function getOAuthTokenEndpoint(config = getCloudConfig()): string {
  const tenant = process.env.EWS_TENANT_ID?.trim() || 'common';
  return `${config.authorityHost}/${tenant}/oauth2/v2.0/token`;
}

export function getEwsTokenScopes(config = getCloudConfig()): string[] {
  return [
    `${config.ewsResource}/EWS.AccessAsUser.All offline_access`,
    `${config.ewsResource}/.default offline_access`,
  ];
}
