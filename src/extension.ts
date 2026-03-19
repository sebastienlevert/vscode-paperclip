import * as vscode from "vscode";
import {
  isWorkspaceInOneDrive,
  isInOneDrive,
  discoverOneDriveRoots,
  findOneDriveRoot,
} from "./onedrive";
import { shareFile, openInOfficeApp, openOnWeb, openFolderOnWeb, openVersionHistory } from "./sharing";
import { OfficePreviewProvider } from "./preview";
import { SyncStatusDecorationProvider } from "./sync-decorations";
import { showRecentFiles } from "./recent-files";

export function activate(context: vscode.ExtensionContext): void {
  if (process.platform !== "win32") {
    return; // Windows only
  }

  // Set initial context for menu visibility
  refreshOneDriveContext();

  context.subscriptions.push(
    vscode.workspace.onDidChangeWorkspaceFolders(() =>
      refreshOneDriveContext()
    )
  );

  // ── Custom editor ────────────────────────────────────────
  context.subscriptions.push(OfficePreviewProvider.register(context));

  // ── Sync status decorations ────────────────────────────
  SyncStatusDecorationProvider.register(context);

  // ── Commands ──────────────────────────────────────────────

  context.subscriptions.push(
    vscode.commands.registerCommand(
      "paperclipped.share",
      async (uri?: vscode.Uri) => {
        const filePath = resolveFilePath(uri);
        if (!filePath) {
          return;
        }
        if (!isInOneDrive(filePath)) {
          vscode.window.showWarningMessage(
            "This file is not in a OneDrive folder."
          );
          return;
        }
        await shareFile(filePath);
      }
    ),

    vscode.commands.registerCommand(
      "paperclipped.openInWord",
      async (uri?: vscode.Uri) => {
        const filePath = resolveFilePath(uri);
        if (filePath) {
          await openInOfficeApp(filePath, "word");
        }
      }
    ),

    vscode.commands.registerCommand(
      "paperclipped.openInExcel",
      async (uri?: vscode.Uri) => {
        const filePath = resolveFilePath(uri);
        if (filePath) {
          await openInOfficeApp(filePath, "excel");
        }
      }
    ),

    vscode.commands.registerCommand(
      "paperclipped.openInPowerPoint",
      async (uri?: vscode.Uri) => {
        const filePath = resolveFilePath(uri);
        if (filePath) {
          await openInOfficeApp(filePath, "powerpoint");
        }
      }
    ),

    vscode.commands.registerCommand(
      "paperclipped.openOnWeb",
      async (uri?: vscode.Uri) => {
        const filePath = resolveFilePath(uri);
        if (!filePath) {
          return;
        }
        if (!isInOneDrive(filePath)) {
          vscode.window.showWarningMessage(
            "This file is not in a OneDrive folder."
          );
          return;
        }
        await openOnWeb(filePath);
      }
    ),

    vscode.commands.registerCommand(
      "paperclipped.openFolderOnWeb",
      async (uri?: vscode.Uri) => {
        const filePath = resolveFilePath(uri);
        if (!filePath) {
          return;
        }
        if (!isInOneDrive(filePath)) {
          vscode.window.showWarningMessage(
            "This file is not in a OneDrive folder."
          );
          return;
        }
        await openFolderOnWeb(filePath);
      }
    ),

    vscode.commands.registerCommand(
      "paperclipped.versionHistory",
      async (uri?: vscode.Uri) => {
        const filePath = resolveFilePath(uri);
        if (!filePath) {
          return;
        }
        if (!isInOneDrive(filePath)) {
          vscode.window.showWarningMessage(
            "This file is not in a OneDrive folder."
          );
          return;
        }
        await openVersionHistory(filePath);
      }
    ),

    vscode.commands.registerCommand(
      "paperclipped.recentFiles",
      async () => {
        await showRecentFiles();
      }
    )
  );

  // ── Status bar ────────────────────────────────────────────
  const statusBarItem = vscode.window.createStatusBarItem(
    vscode.StatusBarAlignment.Right,
    50
  );
  statusBarItem.command = "paperclipped.quickActions";
  context.subscriptions.push(statusBarItem);

  const updateStatusBar = () => {
    const filePath = getActiveFilePath();
    if (filePath && isInOneDrive(filePath)) {
      const root = findOneDriveRoot(filePath);
      const label = root?.accountName ?? "OneDrive";
      const type =
        root?.accountType === "business"
          ? "Work or School"
          : root?.accountType === "personal"
            ? "Personal"
            : "";
      statusBarItem.text = `$(cloud) ${label}`;
      statusBarItem.tooltip = type
        ? `Paperclipped — ${label} (${type})`
        : `Paperclipped — ${label}`;
      statusBarItem.show();
    } else {
      statusBarItem.hide();
    }
  };

  context.subscriptions.push(
    vscode.window.onDidChangeActiveTextEditor(() => updateStatusBar()),
    vscode.window.tabGroups.onDidChangeTabs(() => updateStatusBar())
  );
  updateStatusBar();

  // Quick actions command (triggered by clicking the status bar)
  context.subscriptions.push(
    vscode.commands.registerCommand("paperclipped.quickActions", async () => {
      const filePath = getActiveFilePath();
      if (!filePath || !isInOneDrive(filePath)) {
        return;
      }

      const items: vscode.QuickPickItem[] = [
        { label: "$(globe) Open on Web", description: "Open file in browser" },
        { label: "$(folder) Open Folder on Web", description: "Open parent folder in browser" },
        { label: "$(share) Share", description: "Open sharing dialog" },
        { label: "$(history) Version History", description: "View version history" },
        { label: "$(clock) Recent Files", description: "Browse recent OneDrive files" },
      ];

      const pick = await vscode.window.showQuickPick(items, {
        placeHolder: "Paperclipped — choose an action",
      });

      if (!pick) {
        return;
      }

      const uri = vscode.Uri.file(filePath);
      const commandMap: Record<string, string> = {
        "$(globe) Open on Web": "paperclipped.openOnWeb",
        "$(folder) Open Folder on Web": "paperclipped.openFolderOnWeb",
        "$(share) Share": "paperclipped.share",
        "$(history) Version History": "paperclipped.versionHistory",
        "$(clock) Recent Files": "paperclipped.recentFiles",
      };

      const cmd = commandMap[pick.label];
      if (cmd) {
        vscode.commands.executeCommand(cmd, uri);
      }
    })
  );

  // ── Startup log ──────────────────────────────────────────
  const roots = discoverOneDriveRoots();
  if (roots.length > 0) {
    console.log(
      `[Paperclipped] Detected ${roots.length} root(s): ${roots.map((r) => r.localPath).join(", ")}`
    );
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function refreshOneDriveContext(): void {
  const active = isWorkspaceInOneDrive(vscode.workspace.workspaceFolders);
  vscode.commands.executeCommand(
    "setContext",
    "paperclipped:isOneDriveWorkspace",
    active
  );
}

function resolveFilePath(uri?: vscode.Uri): string | undefined {
  if (uri) {
    return uri.fsPath;
  }

  // Check active text editor
  const editor = vscode.window.activeTextEditor;
  if (editor) {
    return editor.document.uri.fsPath;
  }

  // Check active tab (works for custom editors, non-text files, and when
  // focus is on the sidebar/panel)
  const activeTab = vscode.window.tabGroups.activeTabGroup.activeTab;
  if (activeTab?.input) {
    const input = activeTab.input as any;
    // TabInputText, TabInputCustom, TabInputNotebook all have .uri
    if (input.uri?.fsPath) {
      return input.uri.fsPath;
    }
    // TabInputTextDiff has .modified
    if (input.modified?.fsPath) {
      return input.modified.fsPath;
    }
  }

  // Last resort: check the first visible text editor
  if (vscode.window.visibleTextEditors.length > 0) {
    return vscode.window.visibleTextEditors[0].document.uri.fsPath;
  }

  vscode.window.showWarningMessage("No file selected.");
  return undefined;
}

/** Get the active file path without showing warnings — used for status bar updates. */
function getActiveFilePath(): string | undefined {
  const editor = vscode.window.activeTextEditor;
  if (editor) {
    return editor.document.uri.fsPath;
  }
  const activeTab = vscode.window.tabGroups.activeTabGroup.activeTab;
  if (activeTab?.input) {
    const input = activeTab.input as any;
    if (input.uri?.fsPath) {
      return input.uri.fsPath;
    }
    if (input.modified?.fsPath) {
      return input.modified.fsPath;
    }
  }
  return undefined;
}

export function deactivate(): void {
  // nothing to clean up
}
