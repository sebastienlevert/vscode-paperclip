import * as vscode from "vscode";
import {
  isWorkspaceInOneDrive,
  isInOneDrive,
  discoverOneDriveRoots,
} from "./onedrive";
import { shareFile, openInOfficeApp, openOnWeb } from "./sharing";
import { OfficePreviewProvider } from "./preview";

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
    )
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

export function deactivate(): void {
  // nothing to clean up
}
