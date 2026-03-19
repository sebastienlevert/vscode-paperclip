import * as path from "path";
import * as fs from "fs";
import * as vscode from "vscode";
import {
  findOneDriveRoot,
  enrichAccountsFromRegistry,
  discoverOneDriveRoots,
  buildWebUrl,
} from "./onedrive";

// ---------------------------------------------------------------------------
// Office file type metadata
// ---------------------------------------------------------------------------

interface OfficeFileInfo {
  name: string;
  color: string;
  letter: string;
  commandId: string;
  webAction: string;
}

const OFFICE_MAP: Record<string, OfficeFileInfo> = {
  ".doc": { name: "Word", color: "#2b579a", letter: "W", commandId: "paperclipped.openInWord", webAction: "default" },
  ".docx": { name: "Word", color: "#2b579a", letter: "W", commandId: "paperclipped.openInWord", webAction: "default" },
  ".docm": { name: "Word", color: "#2b579a", letter: "W", commandId: "paperclipped.openInWord", webAction: "default" },
  ".dot": { name: "Word", color: "#2b579a", letter: "W", commandId: "paperclipped.openInWord", webAction: "default" },
  ".dotx": { name: "Word", color: "#2b579a", letter: "W", commandId: "paperclipped.openInWord", webAction: "default" },
  ".dotm": { name: "Word", color: "#2b579a", letter: "W", commandId: "paperclipped.openInWord", webAction: "default" },
  ".rtf": { name: "Word", color: "#2b579a", letter: "W", commandId: "paperclipped.openInWord", webAction: "default" },
  ".xls": { name: "Excel", color: "#217346", letter: "X", commandId: "paperclipped.openInExcel", webAction: "default" },
  ".xlsx": { name: "Excel", color: "#217346", letter: "X", commandId: "paperclipped.openInExcel", webAction: "default" },
  ".xlsm": { name: "Excel", color: "#217346", letter: "X", commandId: "paperclipped.openInExcel", webAction: "default" },
  ".xlsb": { name: "Excel", color: "#217346", letter: "X", commandId: "paperclipped.openInExcel", webAction: "default" },
  ".xlt": { name: "Excel", color: "#217346", letter: "X", commandId: "paperclipped.openInExcel", webAction: "default" },
  ".xltx": { name: "Excel", color: "#217346", letter: "X", commandId: "paperclipped.openInExcel", webAction: "default" },
  ".xltm": { name: "Excel", color: "#217346", letter: "X", commandId: "paperclipped.openInExcel", webAction: "default" },
  ".csv": { name: "Excel", color: "#217346", letter: "X", commandId: "paperclipped.openInExcel", webAction: "default" },
  ".ppt": { name: "PowerPoint", color: "#b7472a", letter: "P", commandId: "paperclipped.openInPowerPoint", webAction: "default" },
  ".pptx": { name: "PowerPoint", color: "#b7472a", letter: "P", commandId: "paperclipped.openInPowerPoint", webAction: "default" },
  ".pptm": { name: "PowerPoint", color: "#b7472a", letter: "P", commandId: "paperclipped.openInPowerPoint", webAction: "default" },
  ".pot": { name: "PowerPoint", color: "#b7472a", letter: "P", commandId: "paperclipped.openInPowerPoint", webAction: "default" },
  ".potx": { name: "PowerPoint", color: "#b7472a", letter: "P", commandId: "paperclipped.openInPowerPoint", webAction: "default" },
  ".potm": { name: "PowerPoint", color: "#b7472a", letter: "P", commandId: "paperclipped.openInPowerPoint", webAction: "default" },
  ".ppsx": { name: "PowerPoint", color: "#b7472a", letter: "P", commandId: "paperclipped.openInPowerPoint", webAction: "default" },
  ".ppsm": { name: "PowerPoint", color: "#b7472a", letter: "P", commandId: "paperclipped.openInPowerPoint", webAction: "default" },
  ".pps": { name: "PowerPoint", color: "#b7472a", letter: "P", commandId: "paperclipped.openInPowerPoint", webAction: "default" },
  ".pdf": { name: "PDF", color: "#d63b2f", letter: "A", commandId: "", webAction: "view" },
};

// ---------------------------------------------------------------------------
// Custom Editor Provider
// ---------------------------------------------------------------------------

export class OfficePreviewProvider implements vscode.CustomReadonlyEditorProvider {
  public static readonly viewType = "paperclipped.officePreview";

  constructor(private readonly context: vscode.ExtensionContext) {}

  public static register(context: vscode.ExtensionContext): vscode.Disposable {
    const provider = new OfficePreviewProvider(context);
    return vscode.window.registerCustomEditorProvider(
      OfficePreviewProvider.viewType,
      provider,
      {
        webviewOptions: { retainContextWhenHidden: true },
        supportsMultipleEditorsPerDocument: false,
      }
    );
  }

  async openCustomDocument(uri: vscode.Uri): Promise<vscode.CustomDocument> {
    return { uri, dispose: () => {} };
  }

  async resolveCustomEditor(
    document: vscode.CustomDocument,
    webviewPanel: vscode.WebviewPanel
  ): Promise<void> {
    const filePath = document.uri.fsPath;
    const ext = path.extname(filePath).toLowerCase();
    const info = OFFICE_MAP[ext];
    const fileName = path.basename(filePath);

    // File metadata
    let fileSize = "";
    let lastModified = "";
    try {
      const stat = fs.statSync(filePath);
      fileSize = formatFileSize(stat.size);
      lastModified = stat.mtime.toLocaleDateString(undefined, {
        year: "numeric",
        month: "short",
        day: "numeric",
        hour: "2-digit",
        minute: "2-digit",
      });
    } catch {
      // ignore stat errors
    }

    // Resolve web URL for iframe
    let embedUrl = "";
    const account = findOneDriveRoot(filePath);
    if (account) {
      if (!account.webEndpoint) {
        await enrichAccountsFromRegistry(discoverOneDriveRoots());
      }
      const webUrl = buildWebUrl(filePath, account);
      if (webUrl) {
        embedUrl = webUrl.replace("action=default", "action=embedview");
      }
    }

    const appName = info?.name ?? "Document";
    const appColor = info?.color ?? "#666666";
    const appLetter = info?.letter ?? "?";
    const hasOpenInApp = info?.commandId ? true : false;

    webviewPanel.webview.options = { enableScripts: true };

    webviewPanel.webview.html = getPreviewHtml({
      fileName,
      fileSize,
      lastModified,
      appName,
      appColor,
      appLetter,
      embedUrl,
      hasOpenInApp,
    });

    // Handle messages from the webview
    webviewPanel.webview.onDidReceiveMessage(async (message) => {
      const uri = document.uri;
      switch (message.command) {
        case "openInApp":
          if (info?.commandId) {
            await vscode.commands.executeCommand(info.commandId, uri);
          }
          break;
        case "openOnWeb":
          await vscode.commands.executeCommand("paperclipped.openOnWeb", uri);
          break;
        case "share":
          await vscode.commands.executeCommand("paperclipped.share", uri);
          break;
      }
    });
  }
}

// ---------------------------------------------------------------------------
// HTML generation
// ---------------------------------------------------------------------------

function formatFileSize(bytes: number): string {
  if (bytes < 1024) { return `${bytes} B`; }
  if (bytes < 1024 * 1024) { return `${(bytes / 1024).toFixed(1)} KB`; }
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

interface PreviewHtmlOptions {
  fileName: string;
  fileSize: string;
  lastModified: string;
  appName: string;
  appColor: string;
  appLetter: string;
  embedUrl: string;
  hasOpenInApp: boolean;
}

function getPreviewHtml(opts: PreviewHtmlOptions): string {
  const iframeSection = opts.embedUrl
    ? `<iframe src="${opts.embedUrl}" frameborder="0" style="width:100%;flex:1;border:none;"></iframe>`
    : `<div class="no-preview">
        <p>Preview not available</p>
        <p class="hint">Use the buttons above to open in the desktop app or on the web.</p>
      </div>`;

  const openInAppButton = opts.hasOpenInApp
    ? `<button class="btn" onclick="send('openInApp')">📄 Open in ${opts.appName}</button>`
    : "";

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="Content-Security-Policy"
    content="default-src 'none';
      style-src 'unsafe-inline';
      script-src 'unsafe-inline';
      frame-src https://*.sharepoint.com https://*.office.com https://*.officeppe.com https://*.officeapps.live.com;">
  <style>
    body {
      margin: 0;
      padding: 0;
      font-family: var(--vscode-font-family, sans-serif);
      color: var(--vscode-foreground);
      background: var(--vscode-editor-background);
      display: flex;
      flex-direction: column;
      height: 100vh;
    }
    .toolbar {
      display: flex;
      align-items: center;
      gap: 12px;
      padding: 10px 16px;
      background: var(--vscode-editorWidget-background, #252526);
      border-bottom: 1px solid var(--vscode-editorWidget-border, #454545);
      flex-shrink: 0;
    }
    .badge {
      width: 36px;
      height: 36px;
      border-radius: 6px;
      display: flex;
      align-items: center;
      justify-content: center;
      font-weight: bold;
      font-size: 18px;
      color: white;
      flex-shrink: 0;
    }
    .file-info {
      flex: 1;
      min-width: 0;
    }
    .file-name {
      font-weight: 600;
      font-size: 13px;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    .file-meta {
      font-size: 11px;
      opacity: 0.7;
      margin-top: 2px;
    }
    .actions {
      display: flex;
      gap: 8px;
      flex-shrink: 0;
    }
    .btn {
      padding: 5px 12px;
      border: 1px solid var(--vscode-button-border, transparent);
      border-radius: 4px;
      background: var(--vscode-button-secondaryBackground, #3a3d41);
      color: var(--vscode-button-secondaryForeground, #cccccc);
      font-size: 12px;
      cursor: pointer;
      white-space: nowrap;
    }
    .btn:hover {
      background: var(--vscode-button-secondaryHoverBackground, #45494e);
    }
    .btn.primary {
      background: var(--vscode-button-background, #0e639c);
      color: var(--vscode-button-foreground, #ffffff);
    }
    .btn.primary:hover {
      background: var(--vscode-button-hoverBackground, #1177bb);
    }
    iframe {
      flex: 1;
    }
    .no-preview {
      flex: 1;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      opacity: 0.6;
    }
    .no-preview p { margin: 4px 0; }
    .no-preview .hint { font-size: 12px; }
  </style>
</head>
<body>
  <div class="toolbar">
    <div class="badge" style="background:${opts.appColor}">${opts.appLetter}</div>
    <div class="file-info">
      <div class="file-name">${escapeHtml(opts.fileName)}</div>
      <div class="file-meta">${opts.appName} · ${opts.fileSize} · ${opts.lastModified}</div>
    </div>
    <div class="actions">
      ${openInAppButton}
      <button class="btn" onclick="send('openOnWeb')">🌐 Open on Web</button>
      <button class="btn primary" onclick="send('share')">📤 Share</button>
    </div>
  </div>
  ${iframeSection}
  <script>
    const vscode = acquireVsCodeApi();
    function send(cmd) { vscode.postMessage({ command: cmd }); }
  </script>
</body>
</html>`;
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
