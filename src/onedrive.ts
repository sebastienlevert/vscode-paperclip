import * as path from "path";
import * as fs from "fs";
import * as os from "os";
import { execFile } from "child_process";

export interface OneDriveAccount {
  localPath: string;
  accountType: "business" | "personal";
  accountName: string;
  webEndpoint?: string;
}

let cachedAccounts: OneDriveAccount[] | undefined;

/**
 * Discover OneDrive root folders from environment variables and the file system.
 */
export function discoverOneDriveRoots(): OneDriveAccount[] {
  if (cachedAccounts) {
    return cachedAccounts;
  }

  const accounts: OneDriveAccount[] = [];
  const seen = new Set<string>();

  const addAccount = (
    localPath: string,
    accountType: "business" | "personal",
    accountName: string
  ) => {
    const key = localPath.toLowerCase();
    if (!seen.has(key) && fs.existsSync(localPath)) {
      seen.add(key);
      accounts.push({ localPath, accountType, accountName });
    }
  };

  // Environment variables
  if (process.env.OneDriveCommercial) {
    addAccount(process.env.OneDriveCommercial, "business", "OneDriveCommercial");
  }
  if (process.env.OneDriveConsumer) {
    addAccount(process.env.OneDriveConsumer, "personal", "OneDriveConsumer");
  }
  if (process.env.OneDrive) {
    addAccount(process.env.OneDrive, "personal", "OneDrive");
  }

  // Scan home directory for OneDrive folders
  const home = os.homedir();
  try {
    for (const entry of fs.readdirSync(home, { withFileTypes: true })) {
      if (entry.isDirectory() && entry.name.startsWith("OneDrive")) {
        const isBusiness = entry.name.includes(" - ");
        addAccount(
          path.join(home, entry.name),
          isBusiness ? "business" : "personal",
          entry.name
        );
      }
    }
  } catch {
    // Ignore errors reading home directory
  }

  cachedAccounts = accounts;
  return accounts;
}

export function findOneDriveRoot(
  filePath: string
): OneDriveAccount | undefined {
  const lowerPath = filePath.toLowerCase();
  let best: OneDriveAccount | undefined;

  for (const account of discoverOneDriveRoots()) {
    const lowerRoot = account.localPath.toLowerCase();
    if (
      lowerPath.startsWith(lowerRoot + path.sep) ||
      lowerPath === lowerRoot
    ) {
      if (!best || account.localPath.length > best.localPath.length) {
        best = account;
      }
    }
  }
  return best;
}

export function isInOneDrive(filePath: string): boolean {
  return findOneDriveRoot(filePath) !== undefined;
}

export function isWorkspaceInOneDrive(
  folders: readonly { uri: { fsPath: string } }[] | undefined
): boolean {
  if (!folders) {
    return false;
  }
  return folders.some((f) => isInOneDrive(f.uri.fsPath));
}

/**
 * Read OneDrive account web endpoints from the Windows registry.
 */
export async function enrichAccountsFromRegistry(
  accounts: OneDriveAccount[]
): Promise<void> {
  const script = `
    $base = 'HKCU:\\Software\\Microsoft\\OneDrive\\Accounts'
    if (Test-Path $base) {
      Get-ChildItem $base | ForEach-Object {
        $props = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
        if ($props.UserFolder -and $props.ServiceEndpointUri) {
          Write-Output "$($props.UserFolder)|$($props.ServiceEndpointUri)"
        }
      }
    }
  `;

  try {
    const output = await runPowerShell(script);
    for (const line of output.split("\n").filter(Boolean)) {
      const [userFolder, serviceUri] = line.trim().split("|");
      if (!userFolder || !serviceUri) {
        continue;
      }
      const match = accounts.find(
        (a) => a.localPath.toLowerCase() === userFolder.toLowerCase()
      );
      if (match) {
        // Strip the /_api suffix so we get a clean base URL
        match.webEndpoint = serviceUri.replace(/\/_api\/?$/, "/");
      }
    }
  } catch {
    // Registry read failed — web features will be limited
  }
}

/**
 * Build a web URL for a file in a OneDrive-synced folder.
 * Falls back to undefined if the web endpoint is not known.
 */
export function buildWebUrl(
  filePath: string,
  account: OneDriveAccount
): string | undefined {
  if (!account.webEndpoint) {
    return undefined;
  }

  const relativePath = path
    .relative(account.localPath, filePath)
    .replace(/\\/g, "/");

  const base = account.webEndpoint.endsWith("/")
    ? account.webEndpoint
    : account.webEndpoint + "/";

  return `${base}Documents/${relativePath}?web=1`;
}

/** Run a PowerShell snippet and return stdout. */
export function runPowerShell(script: string): Promise<string> {
  return new Promise((resolve, reject) => {
    execFile(
      "powershell.exe",
      ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", script],
      { timeout: 15000 },
      (error, stdout, stderr) => {
        if (error) {
          reject(new Error(stderr || error.message));
        } else {
          resolve(stdout);
        }
      }
    );
  });
}

/**
 * Escape a string for embedding inside a PowerShell single-quoted string.
 * Single quotes are doubled: ' → ''
 */
export function psEscape(s: string): string {
  return s.replace(/'/g, "''");
}
