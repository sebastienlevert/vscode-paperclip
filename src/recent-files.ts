import * as path from "path";
import * as fs from "fs";
import * as vscode from "vscode";
import { discoverOneDriveRoots, isInOneDrive } from "./onedrive";

const OFFICE_EXTENSIONS = new Set([
  ".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm", ".rtf",
  ".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx", ".xltm", ".csv",
  ".ppt", ".pptx", ".pptm", ".pot", ".potx", ".potm", ".ppsx", ".ppsm", ".pps",
  ".pdf", ".md",
]);

interface RecentFile {
  filePath: string;
  fileName: string;
  modified: Date;
  ext: string;
}

/**
 * Scan OneDrive roots for recently modified Office and markdown files.
 */
function scanRecentFiles(maxAgeDays: number = 30, limit: number = 50): RecentFile[] {
  const roots = discoverOneDriveRoots();
  const cutoff = new Date(Date.now() - maxAgeDays * 24 * 60 * 60 * 1000);
  const results: RecentFile[] = [];

  for (const root of roots) {
    collectRecent(root.localPath, cutoff, results, limit * 2);
  }

  results.sort((a, b) => b.modified.getTime() - a.modified.getTime());
  return results.slice(0, limit);
}

function collectRecent(
  dir: string,
  cutoff: Date,
  results: RecentFile[],
  limit: number
): void {
  if (results.length >= limit) {
    return;
  }

  let entries: fs.Dirent[];
  try {
    entries = fs.readdirSync(dir, { withFileTypes: true });
  } catch {
    return;
  }

  for (const entry of entries) {
    if (results.length >= limit) {
      return;
    }

    const fullPath = path.join(dir, entry.name);

    if (entry.isDirectory()) {
      // Skip common non-relevant directories
      if (entry.name.startsWith(".") || entry.name === "node_modules") {
        continue;
      }
      collectRecent(fullPath, cutoff, results, limit);
    } else if (entry.isFile()) {
      const ext = path.extname(entry.name).toLowerCase();
      if (!OFFICE_EXTENSIONS.has(ext)) {
        continue;
      }
      // Skip temp files
      if (entry.name.startsWith("~$")) {
        continue;
      }
      try {
        const stat = fs.statSync(fullPath);
        if (stat.mtime >= cutoff) {
          results.push({
            filePath: fullPath,
            fileName: entry.name,
            modified: stat.mtime,
            ext,
          });
        }
      } catch {
        // Skip inaccessible files (e.g., cloud-only)
      }
    }
  }
}

/** Icon for file extension. */
function getIcon(ext: string): string {
  if ([".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm", ".rtf"].includes(ext)) {
    return "$(file-text)";
  }
  if ([".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx", ".xltm", ".csv"].includes(ext)) {
    return "$(table)";
  }
  if ([".ppt", ".pptx", ".pptm", ".pot", ".potx", ".potm", ".ppsx", ".ppsm", ".pps"].includes(ext)) {
    return "$(preview)";
  }
  if (ext === ".pdf") {
    return "$(file-pdf)";
  }
  if (ext === ".md") {
    return "$(markdown)";
  }
  return "$(file)";
}

function formatRelativeTime(date: Date): string {
  const diff = Date.now() - date.getTime();
  const minutes = Math.floor(diff / 60000);
  if (minutes < 60) {
    return `${minutes}m ago`;
  }
  const hours = Math.floor(minutes / 60);
  if (hours < 24) {
    return `${hours}h ago`;
  }
  const days = Math.floor(hours / 24);
  if (days < 7) {
    return `${days}d ago`;
  }
  return date.toLocaleDateString();
}

/**
 * Show a quick pick of recently modified OneDrive files.
 */
export async function showRecentFiles(): Promise<void> {
  const quickPick = vscode.window.createQuickPick();
  quickPick.placeholder = "Search recent OneDrive files...";
  quickPick.busy = true;
  quickPick.show();

  // Scan in background to keep UI responsive
  const files = await new Promise<RecentFile[]>((resolve) => {
    setTimeout(() => resolve(scanRecentFiles()), 0);
  });

  if (files.length === 0) {
    quickPick.items = [{ label: "No recent files found", description: "in OneDrive folders" }];
    quickPick.busy = false;
    return;
  }

  quickPick.items = files.map((f) => ({
    label: `${getIcon(f.ext)} ${f.fileName}`,
    description: formatRelativeTime(f.modified),
    detail: f.filePath,
  }));
  quickPick.busy = false;

  quickPick.onDidAccept(async () => {
    const selected = quickPick.selectedItems[0];
    if (selected?.detail) {
      quickPick.dispose();
      const uri = vscode.Uri.file(selected.detail);
      await vscode.commands.executeCommand("vscode.open", uri);
    }
  });

  quickPick.onDidHide(() => quickPick.dispose());
}
