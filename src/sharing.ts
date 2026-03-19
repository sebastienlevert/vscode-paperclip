import * as path from "path";
import * as vscode from "vscode";
import {
  findOneDriveRoot,
  enrichAccountsFromRegistry,
  discoverOneDriveRoots,
  buildWebUrl,
  runPowerShell,
  psEscape,
} from "./onedrive";

// ---------------------------------------------------------------------------
// Share
// ---------------------------------------------------------------------------

/**
 * Open the native Windows / OneDrive sharing dialog for a file.
 *
 * Uses the Shell.Application COM object to invoke the "Share" verb that the
 * OneDrive sync client adds to the explorer context menu.  Falls back to
 * opening File Explorer with the file selected if the verb is not found.
 */
export async function shareFile(filePath: string): Promise<void> {
  const folderPath = psEscape(path.dirname(filePath));
  const fileName = psEscape(path.basename(filePath));

  const script = `
    $shell = New-Object -ComObject Shell.Application
    $folder = $shell.NameSpace('${folderPath}')
    if (-not $folder) { Write-Output 'FOLDER_NOT_FOUND'; exit 1 }
    $item = $folder.ParseName('${fileName}')
    if (-not $item) { Write-Output 'FILE_NOT_FOUND'; exit 1 }

    $shareVerb = $null
    foreach ($v in $item.Verbs()) {
      if ($v.Name -match '[Ss]hare|[Pp]artag') {
        $shareVerb = $v
        break
      }
    }

    if ($shareVerb) {
      $shareVerb.DoIt()
      Write-Output 'OK'
    } else {
      Write-Output 'NO_SHARE_VERB'
    }
  `;

  try {
    const result = (await runPowerShell(script)).trim();

    if (result === "NO_SHARE_VERB") {
      // Fallback: open File Explorer with the file selected
      await runPowerShell(
        `Start-Process explorer.exe -ArgumentList '/select,"${psEscape(filePath)}"'`
      );
      vscode.window.showInformationMessage(
        "OneDrive share dialog not available. File Explorer opened — right-click the file to share."
      );
    } else if (result !== "OK") {
      throw new Error(result);
    }
  } catch (err: any) {
    vscode.window.showErrorMessage(
      `Failed to open share dialog: ${err.message}`
    );
  }
}

// ---------------------------------------------------------------------------
// Open in Office app
// ---------------------------------------------------------------------------

type OfficeApp = "word" | "excel" | "powerpoint";

const OFFICE_REGISTRY_KEYS: Record<OfficeApp, string> = {
  word: "WINWORD.EXE",
  excel: "EXCEL.EXE",
  powerpoint: "POWERPNT.EXE",
};

/**
 * Open a file in the specified Office desktop application.
 *
 * Looks up the app path from the Windows registry (App Paths) so we always
 * launch the real desktop app rather than whatever the default file handler
 * happens to be.  Falls back to `Start-Process` on the file path itself.
 */
export async function openInOfficeApp(
  filePath: string,
  app: OfficeApp
): Promise<void> {
  const registryKey = OFFICE_REGISTRY_KEYS[app];
  const escaped = psEscape(filePath);

  const script = `
    $regPath = "HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\${registryKey}"
    $appPath = $null
    if (Test-Path $regPath) {
      $appPath = (Get-ItemProperty $regPath -ErrorAction SilentlyContinue).'(default)'
    }
    if ($appPath -and (Test-Path $appPath)) {
      Start-Process -FilePath $appPath -ArgumentList ('"' + '${escaped}' + '"')
      Write-Output 'OK'
    } else {
      Start-Process -FilePath '${escaped}'
      Write-Output 'FALLBACK'
    }
  `;

  try {
    const result = (await runPowerShell(script)).trim();
    if (result === "FALLBACK") {
      const label =
        app.charAt(0).toUpperCase() + app.slice(1);
      vscode.window.showInformationMessage(
        `${label} desktop app not found in registry. Opened with default handler.`
      );
    }
  } catch (err: any) {
    vscode.window.showErrorMessage(`Failed to open in ${app}: ${err.message}`);
  }
}

// ---------------------------------------------------------------------------
// Open on Web
// ---------------------------------------------------------------------------

/**
 * Open a OneDrive-synced file in the browser.
 *
 * Strategy:
 *  1. Build a SharePoint web URL from the registry-based web endpoint.
 *  2. If the endpoint is not known, fall back to the shell "View online" verb.
 */
export async function openOnWeb(filePath: string): Promise<void> {
  const account = findOneDriveRoot(filePath);
  if (!account) {
    vscode.window.showErrorMessage("This file is not in a OneDrive folder.");
    return;
  }

  // Ensure we have the web endpoint
  if (!account.webEndpoint) {
    await enrichAccountsFromRegistry(discoverOneDriveRoots());
  }

  const webUrl = buildWebUrl(filePath, account);
  if (webUrl) {
    await vscode.env.openExternal(vscode.Uri.parse(webUrl));
    return;
  }

  // Fallback: invoke the shell "View online" verb
  const folderPath = psEscape(path.dirname(filePath));
  const fileName = psEscape(path.basename(filePath));

  const script = `
    $shell = New-Object -ComObject Shell.Application
    $folder = $shell.NameSpace('${folderPath}')
    if (-not $folder) { Write-Output 'FOLDER_NOT_FOUND'; exit 1 }
    $item = $folder.ParseName('${fileName}')
    if (-not $item) { Write-Output 'FILE_NOT_FOUND'; exit 1 }

    $webVerb = $null
    foreach ($v in $item.Verbs()) {
      if ($v.Name -match '[Vv]iew.*(online|web)|[Oo]pen.*(browser|web)') {
        $webVerb = $v
        break
      }
    }

    if ($webVerb) {
      $webVerb.DoIt()
      Write-Output 'OK'
    } else {
      Write-Output 'NO_WEB_VERB'
    }
  `;

  try {
    const result = (await runPowerShell(script)).trim();
    if (result === "NO_WEB_VERB") {
      vscode.window.showWarningMessage(
        "Could not determine the web URL for this file. Make sure OneDrive is syncing this folder."
      );
    }
  } catch (err: any) {
    vscode.window.showErrorMessage(
      `Failed to open on web: ${err.message}`
    );
  }
}
