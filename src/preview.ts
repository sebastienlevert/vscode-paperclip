import * as path from "path";
import * as fs from "fs";
import * as vscode from "vscode";
import { findOneDriveRoot, isInOneDrive } from "./onedrive";

// ---------------------------------------------------------------------------
// SVG icon path data (Fluent UI 16×16 paths)
// ---------------------------------------------------------------------------

const WORD_DOC_PATH =
  "M8 4.5V1H4.5C3.67157 1 3 1.67157 3 2.5V13.5C3 14.3284 3.67157 15 4.5 15H11.5C12.3284 15 13 14.3284 13 13.5V6H9.5C8.67157 6 8 5.32843 8 4.5ZM9 4.5V1.25L12.75 5H9.5C9.22386 5 9 4.77614 9 4.5ZM5.5 8H10.5C10.7761 8 11 8.22386 11 8.5C11 8.77614 10.7761 9 10.5 9H5.5C5.22386 9 5 8.77614 5 8.5C5 8.22386 5.22386 8 5.5 8ZM5 10.5C5 10.2239 5.22386 10 5.5 10H10.5C10.7761 10 11 10.2239 11 10.5C11 10.7761 10.7761 11 10.5 11H5.5C5.22386 11 5 10.7761 5 10.5ZM5.5 12H10.5C10.7761 12 11 12.2239 11 12.5C11 12.7761 10.7761 13 10.5 13H5.5C5.22386 13 5 12.7761 5 12.5C5 12.2239 5.22386 12 5.5 12Z";

const EXCEL_GRID_PATH =
  "M4.5 2C3.11929 2 2 3.11929 2 4.5V5H5V2H4.5ZM6 2V5L10 5V2H6ZM5 6H2V10H5V6ZM6 10V6L10 6V10L6 10ZM5 11H2V11.5C2 12.8807 3.11929 14 4.5 14H5V11ZM6 14H10V11L6 11V14ZM11 14V11H14V11.5C14 12.8807 12.8807 14 11.5 14H11ZM14 6V10H11V6H14ZM14 5V4.5C14 3.11929 12.8807 2 11.5 2H11V5H14Z";

const PPT_SLIDE_PATH =
  "M1 5C1 3.89543 1.89543 3 3 3H13C14.1046 3 15 3.89543 15 5V11C15 12.1046 14.1046 13 13 13H3C1.89543 13 1 12.1046 1 11V5ZM4.5 5C4.22386 5 4 5.22386 4 5.5C4 5.77614 4.22386 6 4.5 6H7.5C7.77614 6 8 5.77614 8 5.5C8 5.22386 7.77614 5 7.5 5H4.5ZM4 7.5C4 7.77614 4.22386 8 4.5 8H10.5C10.7761 8 11 7.77614 11 7.5C11 7.22386 10.7761 7 10.5 7H4.5C4.22386 7 4 7.22386 4 7.5ZM4.5 9C4.22386 9 4 9.22386 4 9.5C4 9.77614 4.22386 10 4.5 10H8.5C8.77614 10 9 9.77614 9 9.5C9 9.22386 8.77614 9 8.5 9H4.5Z";

const PDF_DOC_PATH =
  "M4.5 1C3.67157 1 3 1.67157 3 2.5V13.5C3 14.3284 3.67157 15 4.5 15H11.5C12.3284 15 13 14.3284 13 13.5V6H9.5C8.67157 6 8 5.32843 8 4.5V1H4.5ZM9 1.25V4.5C9 4.77614 9.22386 5 9.5 5H12.75L9 1.25ZM5.5 8H10.5C10.7761 8 11 8.22386 11 8.5C11 8.77614 10.7761 9 10.5 9H5.5C5.22386 9 5 8.77614 5 8.5C5 8.22386 5.22386 8 5.5 8ZM5 10.5C5 10.2239 5.22386 10 5.5 10H10.5C10.7761 10 11 10.2239 11 10.5C11 10.7761 10.7761 11 10.5 11H5.5C5.22386 11 5 10.7761 5 10.5ZM5.5 12H10.5C10.7761 12 11 12.2239 11 12.5C11 12.7761 10.7761 13 10.5 13H5.5C5.22386 13 5 12.7761 5 12.5C5 12.2239 5.22386 12 5.5 12Z";

function makeSvg(pathData: string, fill: string, size: number = 18): string {
  return `<svg width="${size}" height="${size}" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="${pathData}" fill="${fill}"/></svg>`;
}

// ---------------------------------------------------------------------------
// Office file type metadata
// ---------------------------------------------------------------------------

interface OfficeFileInfo {
  name: string;
  color: string;
  letter: string;
  commandId: string;
  iconSvg: string;
}

const OFFICE_MAP: Record<string, OfficeFileInfo> = {
  ".doc":  { name: "Word", color: "#185ABD", letter: "W", commandId: "paperclipped.openInWord", iconSvg: makeSvg(WORD_DOC_PATH, "#185ABD") },
  ".docx": { name: "Word", color: "#185ABD", letter: "W", commandId: "paperclipped.openInWord", iconSvg: makeSvg(WORD_DOC_PATH, "#185ABD") },
  ".docm": { name: "Word", color: "#185ABD", letter: "W", commandId: "paperclipped.openInWord", iconSvg: makeSvg(WORD_DOC_PATH, "#185ABD") },
  ".dot":  { name: "Word", color: "#185ABD", letter: "W", commandId: "paperclipped.openInWord", iconSvg: makeSvg(WORD_DOC_PATH, "#185ABD") },
  ".dotx": { name: "Word", color: "#185ABD", letter: "W", commandId: "paperclipped.openInWord", iconSvg: makeSvg(WORD_DOC_PATH, "#185ABD") },
  ".dotm": { name: "Word", color: "#185ABD", letter: "W", commandId: "paperclipped.openInWord", iconSvg: makeSvg(WORD_DOC_PATH, "#185ABD") },
  ".rtf":  { name: "Word", color: "#185ABD", letter: "W", commandId: "paperclipped.openInWord", iconSvg: makeSvg(WORD_DOC_PATH, "#185ABD") },
  ".xls":  { name: "Excel", color: "#107C41", letter: "X", commandId: "paperclipped.openInExcel", iconSvg: makeSvg(EXCEL_GRID_PATH, "#107C41") },
  ".xlsx": { name: "Excel", color: "#107C41", letter: "X", commandId: "paperclipped.openInExcel", iconSvg: makeSvg(EXCEL_GRID_PATH, "#107C41") },
  ".xlsm": { name: "Excel", color: "#107C41", letter: "X", commandId: "paperclipped.openInExcel", iconSvg: makeSvg(EXCEL_GRID_PATH, "#107C41") },
  ".xlsb": { name: "Excel", color: "#107C41", letter: "X", commandId: "paperclipped.openInExcel", iconSvg: makeSvg(EXCEL_GRID_PATH, "#107C41") },
  ".xlt":  { name: "Excel", color: "#107C41", letter: "X", commandId: "paperclipped.openInExcel", iconSvg: makeSvg(EXCEL_GRID_PATH, "#107C41") },
  ".xltx": { name: "Excel", color: "#107C41", letter: "X", commandId: "paperclipped.openInExcel", iconSvg: makeSvg(EXCEL_GRID_PATH, "#107C41") },
  ".xltm": { name: "Excel", color: "#107C41", letter: "X", commandId: "paperclipped.openInExcel", iconSvg: makeSvg(EXCEL_GRID_PATH, "#107C41") },
  ".csv":  { name: "Excel", color: "#107C41", letter: "X", commandId: "paperclipped.openInExcel", iconSvg: makeSvg(EXCEL_GRID_PATH, "#107C41") },
  ".ppt":  { name: "PowerPoint", color: "#C43E1C", letter: "P", commandId: "paperclipped.openInPowerPoint", iconSvg: makeSvg(PPT_SLIDE_PATH, "#C43E1C") },
  ".pptx": { name: "PowerPoint", color: "#C43E1C", letter: "P", commandId: "paperclipped.openInPowerPoint", iconSvg: makeSvg(PPT_SLIDE_PATH, "#C43E1C") },
  ".pptm": { name: "PowerPoint", color: "#C43E1C", letter: "P", commandId: "paperclipped.openInPowerPoint", iconSvg: makeSvg(PPT_SLIDE_PATH, "#C43E1C") },
  ".pot":  { name: "PowerPoint", color: "#C43E1C", letter: "P", commandId: "paperclipped.openInPowerPoint", iconSvg: makeSvg(PPT_SLIDE_PATH, "#C43E1C") },
  ".potx": { name: "PowerPoint", color: "#C43E1C", letter: "P", commandId: "paperclipped.openInPowerPoint", iconSvg: makeSvg(PPT_SLIDE_PATH, "#C43E1C") },
  ".potm": { name: "PowerPoint", color: "#C43E1C", letter: "P", commandId: "paperclipped.openInPowerPoint", iconSvg: makeSvg(PPT_SLIDE_PATH, "#C43E1C") },
  ".ppsx": { name: "PowerPoint", color: "#C43E1C", letter: "P", commandId: "paperclipped.openInPowerPoint", iconSvg: makeSvg(PPT_SLIDE_PATH, "#C43E1C") },
  ".ppsm": { name: "PowerPoint", color: "#C43E1C", letter: "P", commandId: "paperclipped.openInPowerPoint", iconSvg: makeSvg(PPT_SLIDE_PATH, "#C43E1C") },
  ".pps":  { name: "PowerPoint", color: "#C43E1C", letter: "P", commandId: "paperclipped.openInPowerPoint", iconSvg: makeSvg(PPT_SLIDE_PATH, "#C43E1C") },
  ".pdf":  { name: "PDF", color: "#D83B01", letter: "P", commandId: "", iconSvg: makeSvg(PDF_DOC_PATH, "#D83B01") },
};

function getOfficeFileInfo(ext: string): OfficeFileInfo | undefined {
  return OFFICE_MAP[ext.toLowerCase()];
}

// ---------------------------------------------------------------------------
// Custom Editor Provider
// ---------------------------------------------------------------------------

export class OfficePreviewProvider implements vscode.CustomReadonlyEditorProvider {
  public static readonly viewType = "paperclipped.officePreview";

  constructor(private readonly extensionUri: vscode.Uri) {}

  public static register(context: vscode.ExtensionContext): vscode.Disposable {
    const provider = new OfficePreviewProvider(context.extensionUri);
    return vscode.window.registerCustomEditorProvider(
      OfficePreviewProvider.viewType,
      provider,
      {
        webviewOptions: { retainContextWhenHidden: true },
        supportsMultipleEditorsPerDocument: false,
      }
    );
  }

  openCustomDocument(uri: vscode.Uri): vscode.CustomDocument {
    return { uri, dispose() {} };
  }

  async resolveCustomEditor(
    document: vscode.CustomDocument,
    webviewPanel: vscode.WebviewPanel
  ): Promise<void> {
    const filePath = document.uri.fsPath;
    const ext = path.extname(filePath).toLowerCase();
    const fileName = path.basename(filePath);
    const officeInfo = getOfficeFileInfo(ext);
    const inOneDrive = isInOneDrive(filePath);

    // File metadata
    let fileSize = "";
    let modifiedDate = "";
    let createdDate = "";
    try {
      const stat = fs.statSync(filePath);
      fileSize = formatFileSize(stat.size);
      modifiedDate = stat.mtime.toLocaleString();
      createdDate = stat.birthtime.toLocaleString();
    } catch {
      // ignore stat errors
    }

    // OneDrive account info
    const root = findOneDriveRoot(filePath);
    const accountName = root?.accountName ?? "";
    const accountTypeLabel =
      root?.accountType === "business"
        ? "Work or School"
        : root?.accountType === "personal"
          ? "Personal"
          : "";

    webviewPanel.webview.options = { enableScripts: true };
    webviewPanel.webview.html = getPreviewHtml(
      fileName,
      filePath,
      fileSize,
      modifiedDate,
      createdDate,
      accountName,
      accountTypeLabel,
      inOneDrive,
      officeInfo
    );

    // Handle messages from the webview
    webviewPanel.webview.onDidReceiveMessage(async (message) => {
      const uri = vscode.Uri.file(filePath);
      switch (message.command) {
        case "openInApp":
          if (officeInfo) {
            vscode.commands.executeCommand(officeInfo.commandId, uri);
          }
          break;
        case "openOnWeb":
          vscode.commands.executeCommand("paperclipped.openOnWeb", uri);
          break;
        case "share":
          vscode.commands.executeCommand("paperclipped.share", uri);
          break;
      }
    });
  }
}

// ---------------------------------------------------------------------------
// HTML generation
// ---------------------------------------------------------------------------

function formatFileSize(bytes: number): string {
  if (bytes < 1024) {
    return `${bytes} B`;
  }
  if (bytes < 1024 * 1024) {
    return `${(bytes / 1024).toFixed(1)} KB`;
  }
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function getPreviewHtml(
  fileName: string,
  filePath: string,
  fileSize: string,
  modifiedDate: string,
  createdDate: string,
  accountName: string,
  accountTypeLabel: string,
  inOneDrive: boolean,
  officeInfo: OfficeFileInfo | undefined
): string {
  const appName = officeInfo?.name ?? "Office";
  const appColor = officeInfo?.color ?? "#0078D4";
  const appLetter = officeInfo?.letter ?? "?";

  // Map app names to their SVG path data for the hero icon
  const heroPathMap: Record<string, string> = {
    Word: WORD_DOC_PATH,
    Excel: EXCEL_GRID_PATH,
    PowerPoint: PPT_SLIDE_PATH,
    PDF: PDF_DOC_PATH,
  };
  const heroPath = officeInfo ? heroPathMap[officeInfo.name] ?? "" : "";
  const heroContent = heroPath
    ? `<svg width="56" height="56" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="${heroPath}" fill="white"/></svg>`
    : appLetter;

  const nonce = generateNonce();
  const csp = [
    "default-src 'none'",
    `style-src 'nonce-${nonce}'`,
    `script-src 'nonce-${nonce}'`,
  ].join("; ");

  // Build detail rows
  const rows: string[] = [];
  rows.push(detailRow("File name", fileName));
  rows.push(
    detailRow("Type", `${appName} document (${escapeHtml(path.extname(filePath))})`)
  );
  if (fileSize) {
    rows.push(detailRow("Size", fileSize));
  }
  if (modifiedDate) {
    rows.push(detailRow("Modified", modifiedDate));
  }
  if (createdDate) {
    rows.push(detailRow("Created", createdDate));
  }
  rows.push(detailRow("Location", escapeHtml(path.dirname(filePath))));
  if (inOneDrive && accountName) {
    rows.push(
      detailRow(
        "OneDrive",
        `${escapeHtml(accountName)}${accountTypeLabel ? ` (${escapeHtml(accountTypeLabel)})` : ""}`
      )
    );
    rows.push(
      detailRow("Sync status", '<span class="sync-badge">\u25CF Synced</span>')
    );
  }

  // Globe icon (Open on Web)
  const globePath =
    "M6 8C6 7.29718 6.04415 6.62474 6.12456 6H9.87544C9.95585 6.62474 10 7.29718 10 8C10 8.70282 9.95585 9.37526 9.87544 10H6.12456C6.04415 9.37526 6 8.70282 6 8ZM5.11686 10C5.0406 9.36521 5 8.69337 5 8C5 7.30663 5.0406 6.63479 5.11686 6H2.34141C2.12031 6.62556 2 7.29873 2 8C2 8.70127 2.12031 9.37444 2.34141 10H5.11686ZM2.80269 11H5.27206C5.39817 11.6551 5.56493 12.254 5.76556 12.7757C5.89989 13.125 6.05249 13.4476 6.22341 13.7326C4.76902 13.2824 3.55119 12.2939 2.80269 11ZM6.292 11H9.708C9.59779 11.5266 9.46003 12.0035 9.30109 12.4167C9.08782 12.9712 8.84611 13.3857 8.60319 13.6528C8.3604 13.9198 8.15584 14 8 14C7.84416 14 7.6396 13.9198 7.39681 13.6528C7.15389 13.3857 6.91218 12.9712 6.69891 12.4167C6.53997 12.0035 6.40221 11.5266 6.292 11ZM10.7279 11C10.6018 11.6551 10.4351 12.254 10.2344 12.7757C10.1001 13.125 9.94751 13.4476 9.77659 13.7326C11.231 13.2824 12.4488 12.2939 13.1973 11H10.7279ZM13.6586 10C13.8797 9.37444 14 8.70127 14 8C14 7.29873 13.8797 6.62556 13.6586 6H10.8831C10.9594 6.63479 11 7.30663 11 8C11 8.69337 10.9594 9.36521 10.8831 10H13.6586ZM9.30109 3.5833C9.46003 3.99654 9.59779 4.47343 9.708 5H6.292C6.40221 4.47343 6.53997 3.99654 6.69891 3.5833C6.91218 3.02877 7.15389 2.61433 7.39681 2.34719C7.6396 2.08019 7.84416 2 8 2C8.15584 2 8.3604 2.08019 8.60319 2.34719C8.84611 2.61433 9.08782 3.02877 9.30109 3.5833ZM10.7279 5H13.1973C12.4488 3.70607 11.231 2.7176 9.77658 2.26738C9.94751 2.55238 10.1001 2.87505 10.2344 3.22432C10.4351 3.74596 10.6018 4.34494 10.7279 5ZM2.80269 5H5.27206C5.39817 4.34494 5.56493 3.74596 5.76556 3.22432C5.89989 2.87505 6.05249 2.55238 6.22341 2.26738C4.76902 2.7176 3.55119 3.70607 2.80269 5Z";

  // Share / external-link icon
  const sharePath =
    "M7.5 2C7.77614 2 8 2.22386 8 2.5C7.99995 2.77609 7.77611 3 7.5 3H4.5C3.67157 3 3 3.67157 3 4.5V11.5C3.00006 12.3284 3.67161 13 4.5 13H11.5C12.3284 13 12.9999 12.3284 13 11.5V9.5C13 9.22385 13.2239 9 13.5 9C13.7761 9 14 9.22385 14 9.5V11.5C13.9999 12.8807 12.8807 14 11.5 14H4.5C3.11932 14 2.00006 12.8807 2 11.5V4.5C2 3.11929 3.11929 2 4.5 2H7.5ZM10.7803 1.05078C10.9518 0.966903 11.1559 0.988318 11.3066 1.10547L15.8066 4.60547C15.9284 4.70019 16 4.84572 16 5C16 5.15425 15.9284 5.29982 15.8066 5.39453L11.3066 8.89453C11.1559 9.01163 10.9517 9.03306 10.7803 8.94922C10.6088 8.86533 10.5 8.69092 10.5 8.5V7.02539C8.26802 7.25505 6.87621 8.99835 6.10352 10.4238L5.94727 10.7236C5.84354 10.931 5.61041 11.0396 5.38477 10.9863C5.15933 10.9329 5.00005 10.7317 5 10.5C5 8.42355 5.51821 6.55894 6.53711 5.20019C7.47545 3.94908 8.82277 3.15337 10.5 3.02148V1.5C10.5 1.30905 10.6088 1.13469 10.7803 1.05078Z";

  const globeSvg = `<svg width="18" height="18" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="${globePath}" fill="currentColor"/></svg>`;
  const shareSvg = `<svg width="18" height="18" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="${sharePath}" fill="currentColor"/></svg>`;

  // Action buttons
  const hasOpenInApp = !!officeInfo?.commandId;
  const openInAppIcon =
    officeInfo && hasOpenInApp
      ? makeSvg(heroPathMap[officeInfo.name] ?? WORD_DOC_PATH, "#fff")
      : "";

  const openInAppBtn = hasOpenInApp
    ? `<button class="btn btn-primary btn-large" data-action="openInApp">
        <span class="btn-icon">${openInAppIcon}</span>Open in ${escapeHtml(appName)}
      </button>`
    : "";

  const openOnWebBtn = inOneDrive
    ? `<button class="btn btn-secondary btn-large" data-action="openOnWeb">
        <span class="btn-icon">${globeSvg}</span>Open on Web
      </button>`
    : "";

  const shareBtn = inOneDrive
    ? `<button class="btn btn-secondary btn-large" data-action="share">
        <span class="btn-icon">${shareSvg}</span>Share
      </button>`
    : "";

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="Content-Security-Policy" content="${csp}">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style nonce="${nonce}">
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
      font-family: var(--vscode-font-family, 'Segoe UI', sans-serif);
      background: var(--vscode-editor-background);
      color: var(--vscode-editor-foreground);
      height: 100vh;
      display: flex;
      align-items: flex-start;
      justify-content: center;
      overflow-y: auto;
    }

    .container {
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 32px;
      max-width: 480px;
      width: 100%;
      padding: 40px 24px;
      overflow-y: auto;
    }

    /* ── App icon ──────────────────────────── */
    .app-icon {
      width: 120px; height: 120px;
      border-radius: 24px;
      background: ${appColor};
      display: flex; align-items: center; justify-content: center;
      font-size: 56px; font-weight: 700; color: #fff;
      box-shadow: 0 6px 32px rgba(0,0,0,0.22);
    }

    /* ── File info card ──────────────────── */
    .file-card {
      width: 100%;
      border: 1px solid var(--vscode-panel-border, #333);
      border-radius: 8px;
      overflow: hidden;
    }
    .file-card-header {
      padding: 14px 16px;
      font-size: 15px;
      font-weight: 600;
      border-bottom: 1px solid var(--vscode-panel-border, #333);
      background: var(--vscode-sideBar-background, rgba(255,255,255,0.03));
    }
    .file-card-body {
      padding: 0;
    }
    .detail-row {
      display: flex;
      padding: 8px 16px;
      border-bottom: 1px solid var(--vscode-panel-border, rgba(255,255,255,0.06));
      font-size: 13px;
    }
    .detail-row:last-child { border-bottom: none; }
    .detail-label {
      width: 110px;
      flex-shrink: 0;
      color: var(--vscode-descriptionForeground, #888);
      font-weight: 500;
    }
    .detail-value {
      flex: 1;
      min-width: 0;
      word-break: break-all;
    }

    .sync-badge {
      color: #3fb950;
      font-weight: 500;
    }

    /* ── Action buttons ──────────────────── */
    .actions {
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      justify-content: center;
      width: 100%;
    }
    .btn {
      padding: 0;
      border: 1px solid var(--vscode-button-border, transparent);
      border-radius: 6px;
      font-size: 13px;
      cursor: pointer;
      font-family: inherit;
      display: inline-flex;
      align-items: center;
      gap: 8px;
      transition: opacity 0.15s;
    }
    .btn:hover { opacity: 0.88; }
    .btn-large { padding: 10px 20px; font-size: 13px; font-weight: 500; }
    .btn-icon { font-size: 16px; display: inline-flex; align-items: center; }
    .btn-icon svg { vertical-align: middle; }
    .btn-primary {
      background: ${appColor};
      color: #fff;
      border-color: ${appColor};
    }
    .btn-secondary {
      background: var(--vscode-button-secondaryBackground, #333);
      color: var(--vscode-button-secondaryForeground, #fff);
    }

  </style>
</head>
<body>
  <div class="container">
    <div class="app-icon">${heroContent || appLetter}</div>

    <div class="file-card">
      <div class="file-card-header">${escapeHtml(fileName)}</div>
      <div class="file-card-body">
        ${rows.join("\n        ")}
      </div>
    </div>

    <div class="actions">
      ${openInAppBtn}
      ${openOnWebBtn}
      ${shareBtn}
    </div>

  </div>

  <script nonce="${nonce}">
    const vscode = acquireVsCodeApi();
    document.querySelectorAll('[data-action]').forEach(btn => {
      btn.addEventListener('click', () => {
        vscode.postMessage({ command: btn.getAttribute('data-action') });
      });
    });
  </script>
</body>
</html>`;
}

function detailRow(label: string, value: string): string {
  return `<div class="detail-row"><span class="detail-label">${escapeHtml(label)}</span><span class="detail-value">${value}</span></div>`;
}

function generateNonce(): string {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  let nonce = "";
  for (let i = 0; i < 32; i++) {
    nonce += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return nonce;
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
