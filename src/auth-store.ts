import * as fs from 'fs';
import * as path from 'path';
import { spawnSync } from 'child_process';

const DEFAULT_AUTH_STORE_FILE = 'auth.secure.json';
const DPAPI_SCHEME = 'dpapi-v1';

type SecureAuthStoreFile = {
  version: 1;
  scheme: typeof DPAPI_SCHEME;
  clientId: string;
  clientSecret: string;
  redirectUri: string;
  encryptedRefreshToken: string;
  updatedAt: string;
};

export type ResolvedAuthData = {
  clientId: string;
  clientSecret: string;
  redirectUri: string;
  refreshToken: string;
};

export function resolveAuthStorePath(explicitPath?: string): string {
  if (explicitPath && explicitPath.trim()) {
    return path.resolve(explicitPath.trim());
  }
  return path.join(process.cwd(), DEFAULT_AUTH_STORE_FILE);
}

function runPowerShell(script: string, input: string): string {
  const result = spawnSync(
    'powershell',
    ['-NoProfile', '-NonInteractive', '-Command', script],
    { input, encoding: 'utf8' }
  );

  if (result.error) {
    throw result.error;
  }

  if (result.status !== 0) {
    throw new Error(result.stderr?.trim() || 'PowerShell DPAPI call failed');
  }

  return (result.stdout || '').trim();
}

function encryptWithDpapi(plainText: string): string {
  const script = `
$data = [Console]::In.ReadToEnd()
$secure = ConvertTo-SecureString -String $data -AsPlainText -Force
ConvertFrom-SecureString -SecureString $secure
`;
  return runPowerShell(script, plainText);
}

function decryptWithDpapi(cipherTextBase64: string): string {
  const script = `
$data = [Console]::In.ReadToEnd()
$secure = ConvertTo-SecureString -String $data.Trim()
$bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
try {
  [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
} finally {
  [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
}
`;
  return runPowerShell(script, cipherTextBase64);
}

export function saveSecureAuthStore(
  storePath: string,
  data: { clientId: string; clientSecret: string; redirectUri: string; refreshToken: string }
): void {
  if (process.platform !== 'win32') {
    throw new Error('Secure auth store is currently implemented for Windows only.');
  }

  const encryptedRefreshToken = encryptWithDpapi(data.refreshToken);
  const payload: SecureAuthStoreFile = {
    version: 1,
    scheme: DPAPI_SCHEME,
    clientId: data.clientId,
    clientSecret: data.clientSecret,
    redirectUri: data.redirectUri,
    encryptedRefreshToken,
    updatedAt: new Date().toISOString(),
  };

  fs.writeFileSync(storePath, JSON.stringify(payload, null, 2), 'utf8');
}

export function loadSecureAuthStore(storePath: string): ResolvedAuthData | null {
  if (!fs.existsSync(storePath)) {
    return null;
  }

  const raw = fs.readFileSync(storePath, 'utf8');
  const parsed = JSON.parse(raw) as Partial<SecureAuthStoreFile>;

  if (
    parsed.version !== 1 ||
    parsed.scheme !== DPAPI_SCHEME ||
    !parsed.clientId ||
    !parsed.clientSecret ||
    !parsed.redirectUri ||
    !parsed.encryptedRefreshToken
  ) {
    throw new Error('Invalid secure auth store format.');
  }

  if (process.platform !== 'win32') {
    throw new Error('Secure auth store is currently implemented for Windows only.');
  }

  const refreshToken = decryptWithDpapi(parsed.encryptedRefreshToken);

  return {
    clientId: parsed.clientId,
    clientSecret: parsed.clientSecret,
    redirectUri: parsed.redirectUri,
    refreshToken,
  };
}
