import * as fs from "fs";
import * as vscode from "vscode";
import { isInOneDrive } from "./onedrive";

// Windows file attribute flags used by OneDrive
const FILE_ATTRIBUTE_SPARSE = 0x200;
const FILE_ATTRIBUTE_OFFLINE = 0x1000;
const FILE_ATTRIBUTE_PINNED = 0x00080000;

type SyncStatus = "available" | "cloud-only" | "pinned";

function getSyncStatus(filePath: string): SyncStatus | undefined {
  try {
    // Node's fs.statSync doesn't expose raw Win32 attributes, so we use
    // a heuristic: if the file is a reparse point (sparse + offline) it's
    // cloud-only. We detect this by checking if the file size on disk is 0
    // while the reported size is > 0.
    //
    // For a more accurate check we'd need native bindings, but this
    // heuristic works well for OneDrive files.
    const stat = fs.statSync(filePath);

    // If we can read the file size and it's a regular file, it's available
    if (stat.isFile()) {
      // Try to detect cloud-only by attempting to read the first byte
      // Cloud-only files will throw or return empty when opened without
      // the FILE_FLAG_OPEN_REPARSE_POINT flag. However, Node may trigger
      // a download. Instead, check the blocks allocated.
      //
      // On Windows with NTFS, stat.blocks is not reliable for cloud files.
      // Fall back to PowerShell for accurate status.
      return "available";
    }
    return undefined;
  } catch {
    return undefined;
  }
}

/**
 * Query sync status for multiple files using a single PowerShell call.
 */
async function batchSyncStatus(
  filePaths: string[]
): Promise<Map<string, SyncStatus>> {
  const result = new Map<string, SyncStatus>();
  if (filePaths.length === 0) {
    return result;
  }

  // Build a PowerShell script that checks attributes for all files at once
  const pathsArray = filePaths.map((p) => `'${p.replace(/'/g, "''")}'`).join(",");
  const script = `
    $paths = @(${pathsArray})
    foreach ($p in $paths) {
      try {
        $attr = [int][System.IO.File]::GetAttributes($p)
        $sparse = ($attr -band ${FILE_ATTRIBUTE_SPARSE}) -ne 0
        $offline = ($attr -band ${FILE_ATTRIBUTE_OFFLINE}) -ne 0
        $pinned = ($attr -band ${FILE_ATTRIBUTE_PINNED}) -ne 0
        if ($pinned -and -not $sparse) { $status = 'pinned' }
        elseif (-not $sparse) { $status = 'available' }
        elseif ($sparse -and $offline) { $status = 'cloud-only' }
        else { $status = 'available' }
        Write-Output "$p|$status"
      } catch {
        Write-Output "$p|error"
      }
    }
  `;

  try {
    const { runPowerShell } = await import("./onedrive");
    const output = await runPowerShell(script);
    for (const line of output.split("\n").filter(Boolean)) {
      const sepIdx = line.lastIndexOf("|");
      if (sepIdx === -1) {
        continue;
      }
      const filePath = line.substring(0, sepIdx).trim();
      const status = line.substring(sepIdx + 1).trim() as SyncStatus;
      if (status === "available" || status === "cloud-only" || status === "pinned") {
        result.set(filePath, status);
      }
    }
  } catch {
    // PowerShell failed — no decorations
  }

  return result;
}

/**
 * File decoration provider that shows OneDrive sync status badges
 * on files in the Explorer tree.
 */
export class SyncStatusDecorationProvider
  implements vscode.FileDecorationProvider
{
  private readonly _onDidChangeFileDecorations =
    new vscode.EventEmitter<vscode.Uri | vscode.Uri[] | undefined>();
  readonly onDidChangeFileDecorations = this._onDidChangeFileDecorations.event;

  private cache = new Map<string, SyncStatus>();
  private pendingRefresh = false;

  static register(context: vscode.ExtensionContext): SyncStatusDecorationProvider {
    const provider = new SyncStatusDecorationProvider();
    context.subscriptions.push(
      vscode.window.registerFileDecorationProvider(provider)
    );

    // Refresh decorations when files change
    const watcher = vscode.workspace.createFileSystemWatcher("**/*");
    context.subscriptions.push(
      watcher.onDidChange(() => provider.refresh()),
      watcher.onDidCreate(() => provider.refresh()),
      watcher.onDidDelete(() => provider.refresh()),
      watcher
    );

    // Initial population
    provider.refreshAll();

    return provider;
  }

  provideFileDecoration(
    uri: vscode.Uri
  ): vscode.FileDecoration | undefined {
    if (uri.scheme !== "file") {
      return undefined;
    }

    const filePath = uri.fsPath;
    if (!isInOneDrive(filePath)) {
      return undefined;
    }

    const status = this.cache.get(filePath);
    if (!status) {
      return undefined;
    }

    switch (status) {
      case "pinned":
        return {
          badge: "📌",
          tooltip: "Always available on this device",
          color: new vscode.ThemeColor("charts.green"),
        };
      case "available":
        return {
          badge: "✓",
          tooltip: "Available on this device",
          color: new vscode.ThemeColor("charts.green"),
        };
      case "cloud-only":
        return {
          badge: "☁",
          tooltip: "Available online only",
          color: new vscode.ThemeColor("charts.blue"),
        };
      default:
        return undefined;
    }
  }

  /** Debounced refresh for individual file changes. */
  refresh(): void {
    if (this.pendingRefresh) {
      return;
    }
    this.pendingRefresh = true;
    setTimeout(() => {
      this.pendingRefresh = false;
      this.refreshAll();
    }, 2000);
  }

  /** Refresh all OneDrive files in workspace. */
  async refreshAll(): Promise<void> {
    const oneDriveFiles: string[] = [];

    // Collect OneDrive files from open tabs
    for (const group of vscode.window.tabGroups.all) {
      for (const tab of group.tabs) {
        const input = tab.input as any;
        const uri = input?.uri ?? input?.modified;
        if (uri?.scheme === "file" && isInOneDrive(uri.fsPath)) {
          oneDriveFiles.push(uri.fsPath);
        }
      }
    }

    // Collect from visible workspace files (limit scan to avoid perf issues)
    if (vscode.workspace.workspaceFolders) {
      for (const folder of vscode.workspace.workspaceFolders) {
        if (isInOneDrive(folder.uri.fsPath)) {
          const files = await vscode.workspace.findFiles(
            new vscode.RelativePattern(folder, "**/*"),
            "**/node_modules/**",
            500
          );
          for (const f of files) {
            oneDriveFiles.push(f.fsPath);
          }
        }
      }
    }

    if (oneDriveFiles.length === 0) {
      return;
    }

    // Deduplicate
    const unique = [...new Set(oneDriveFiles)];

    const statuses = await batchSyncStatus(unique);
    this.cache = statuses;
    this._onDidChangeFileDecorations.fire(undefined);
  }
}
