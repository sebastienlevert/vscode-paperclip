<p align="center">
  <img src="https://raw.githubusercontent.com/sebastienlevert/vscode-paperclip-office/main/paperclip.png" alt="Paperclip" width="128" />
</p>

# Paperclip

A better integration for VS Code and productivity documents— share files, open on the web, open in Word, Excel, PowerPoint, and preview Office documents without leaving your editor.

> **Windows only** (for now). Requires the OneDrive sync client.

---

## Table of Contents

- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
  - [Sharing files](#sharing-files)
  - [Opening in Office desktop apps](#opening-in-office-desktop-apps)
  - [Opening on the web](#opening-on-the-web)
  - [Office document preview](#office-document-preview)
- [Requirements](#requirements)
- [How it works](#how-it-works)
  - [OneDrive detection](#onedrive-detection)
  - [Share dialog](#share-dialog)
  - [Office app resolution](#office-app-resolution)
  - [Web URL resolution](#web-url-resolution)
  - [Custom editor preview](#custom-editor-preview)
- [Context menu structure](#context-menu-structure)
- [Commands](#commands)
- [Settings & configuration](#settings--configuration)
- [Troubleshooting](#troubleshooting)
- [Known limitations](#known-limitations)
- [Contributing](#contributing)
- [License](#license)

---

## Features

### 📤 Share via OneDrive

Right-click any file in a OneDrive-synced folder → **Paperclip** → **Share…** to open the native Windows OneDrive sharing dialog — the same dialog you see in File Explorer.

### 📄 Open in Office Desktop Apps

Right-click Office documents to open them in their native desktop application:

| Command              | File types                        |
|----------------------|-----------------------------------|
| **Open in Word**     | `.doc`, `.docx`, `.docm`          |
| **Open in Excel**    | `.xls`, `.xlsx`, `.xlsm`, `.csv`  |
| **Open in PowerPoint** | `.ppt`, `.pptx`, `.pptm`        |

The extension resolves the Office executable from the Windows registry (`App Paths`), so it always launches the real desktop app — not the web version or another default handler.

### 🌐 Open on Web

Right-click any file → **Paperclip** → **Open on Web** to view or edit the file in your browser via SharePoint / OneDrive for Business web. Works for all file types, not just Office documents.

### 🔎 Office Document Preview

Double-click an Office file (`.docx`, `.xlsx`, `.pptx`, etc.) to open an in-editor preview. The preview shows:

- A toolbar with an app-colored badge, file name, size, and last-modified date
- Action buttons: **Open in App**, **Open on Web**, **Share**
- An embedded Office Online iframe (when the OneDrive web endpoint is available)

> **Note:** The Office Online iframe requires your browser to be signed in to your SharePoint tenant. If authentication fails inside the webview, use the action buttons as a reliable fallback.

---

## Installation

### From VSIX (local build)

```bash
# Build and package
cd vscode-paperclip-office
npm install
npm run compile
npx vsce package --allow-missing-repository

# Install
code --install-extension vscode-paperclip-office-0.1.0.vsix --force
```

### From source (development)

```bash
git clone <repo-url>
cd vscode-paperclip-office
```

---

## Usage

### Sharing files

1. Open a folder that lives inside a OneDrive-synced directory.
2. In the **Explorer** panel, right-click any file.
3. Select **Paperclip** → **Share…**
4. The native Windows OneDrive sharing dialog opens — pick recipients and permissions as usual.

If the OneDrive share verb is unavailable (e.g., OneDrive is not running), the extension falls back to opening File Explorer with the file selected so you can share from there.

### Opening in Office desktop apps

1. Right-click a Word, Excel, or PowerPoint file in the Explorer.
2. Choose **Paperclip** → **Open in Word** / **Open in Excel** / **Open in PowerPoint**.
3. The file opens in the desktop application.

The extension looks up the app path from the Windows registry:

```
HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE
HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE
HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE
```

If the registry key is not found, it falls back to `Start-Process` on the file, which uses the system default handler.

### Opening on the web

1. Right-click any file in the Explorer (not just Office files).
2. Choose **Paperclip** → **Open on Web**.
3. Your default browser opens the file on SharePoint / OneDrive web.

The extension builds a SharePoint `Doc.aspx` URL from the registry-stored web endpoint. If the endpoint is unavailable, it falls back to the shell "View online" verb.

### Office document preview

1. Double-click any Office file (`.docx`, `.xlsx`, `.pptx`, `.doc`, `.xls`, `.ppt`, etc.).
2. An in-editor preview opens with a toolbar and embedded Office Online view.
3. Use the toolbar buttons to open the file in the desktop app, on the web, or to share it.

To revert to the default VS Code binary editor, right-click the file → **Open With…** → select the built-in editor.

---

## Requirements

| Requirement | Details |
|-------------|---------|
| **Operating system** | Windows 10 or Windows 11 |
| **OneDrive sync client** | Installed and signed in (personal, business, or both) |
| **Office** | Required for "Open in Word/Excel/PowerPoint". The extension gracefully degrades if Office is not installed — it uses the default file handler instead. |
| **VS Code** | 1.85.0 or later |

---

## How It Works

### OneDrive detection

On activation, the extension discovers OneDrive root folders using two strategies:

1. **Environment variables** — `OneDriveCommercial`, `OneDriveConsumer`, `OneDrive`
2. **Home directory scan** — Any directory in `%USERPROFILE%` starting with `OneDrive` (e.g., `OneDrive - Contoso`)

Business accounts are identified by the presence of ` - ` in the folder name. Results are cached for the lifetime of the VS Code window.

When any workspace folder falls inside a discovered OneDrive root, the extension sets the context key `paperclip:isOneDriveWorkspace` to `true`, which activates the context menu contributions.

### Share dialog

Sharing uses the Windows **Shell.Application** COM object:

```
Shell.Application → NameSpace(folder) → ParseName(file) → Verbs() → find "Share" → DoIt()
```

The extension searches for a verb matching `[Ss]hare|[Pp]artag` (supports English and French locales). If the verb isn't found, it falls back to `explorer.exe /select,"<file>"`.

### Office app resolution

The extension reads the real Office executable path from the registry:

```
HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\<exe>
```

For example, Word resolves to `C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE`. The file path is passed via `Start-Process -ArgumentList` with proper double-quoting to handle paths with spaces (e.g., `OneDrive - Contoso`).

### Web URL resolution

The OneDrive sync client stores web endpoints in the registry:

```
HKCU:\Software\Microsoft\OneDrive\Accounts\*
  → UserFolder: C:\Users\user\OneDrive - Contoso
  → ServiceEndpointUri: https://contoso-my.sharepoint.com/personal/user_contoso_com/
```

The extension matches the file's path to a `UserFolder`, computes the relative path, and builds a `Doc.aspx` URL:

```
https://<endpoint>/_layouts/15/Doc.aspx?sourcedoc=<encoded-path>&action=default
```

### Custom editor preview

The `OfficePreviewProvider` implements VS Code's `CustomReadonlyEditorProvider` API. When you open an Office file, it:

1. Resolves the OneDrive account and web endpoint (via registry enrichment)
2. Transforms the `Doc.aspx` URL from `action=default` to `action=embedview`
3. Renders a webview with a toolbar and an `<iframe>` pointing to the embed URL
4. Listens for `postMessage` events from the webview buttons and dispatches the corresponding VS Code commands

The Content Security Policy allows frames from `*.sharepoint.com`, `*.office.com`, `*.officeppe.com`, and `*.officeapps.live.com`.

---

## Context Menu Structure

When you right-click a file in a OneDrive workspace:

```
  ┌─────────────────────────┐
  │ …                       │
  │ Paperclip           ▸   │──┐
  │ …                       │  │  ┌───────────────────────────┐
  └─────────────────────────┘  └──│ 📤 Share…                │
                                  │───────────────────────────│
                                  │ 📘 Open in Word           │ ← .doc/.docx/.docm
                                  │ 📗 Open in Excel          │ ← .xls/.xlsx/.xlsm/.csv
                                  │ 📙 Open in PowerPoint     │ ← .ppt/.pptx/.pptm
                                  │───────────────────────────│
                                  │ 🌐 Open on Web            │
                                  └───────────────────────────┘
```

The **Paperclip** submenu only appears when:
- The workspace folder is inside a OneDrive-synced directory
- The right-clicked item is a file (not a folder)

Office-specific entries (Open in Word/Excel/PowerPoint) are conditionally shown based on the file extension.

---

## Commands

All commands are under the `Paperclip` category and available in the Command Palette (`Ctrl+Shift+P`):

| Command | ID | Description |
|---------|----|-------------|
| **Share…** | `paperclip.share` | Open native OneDrive sharing dialog |
| **Open in Word** | `paperclip.openInWord` | Open `.doc`/`.docx`/`.docm` in Word desktop |
| **Open in Excel** | `paperclip.openInExcel` | Open `.xls`/`.xlsx`/`.xlsm`/`.csv` in Excel desktop |
| **Open in PowerPoint** | `paperclip.openInPowerPoint` | Open `.ppt`/`.pptx`/`.pptm` in PowerPoint desktop |
| **Open on Web** | `paperclip.openOnWeb` | Open file in browser via SharePoint/OneDrive web |

---

## Settings & Configuration

The extension does not currently expose user-configurable settings. Behavior is driven by:

- **OneDrive sync client** state and registry entries
- **Office installation** presence in the Windows registry
- **Workspace folder** location relative to OneDrive roots

---

## Troubleshooting

### "This file is not in a OneDrive folder"

- Make sure your workspace folder is inside a OneDrive-synced directory
- Check that the OneDrive sync client is running and signed in
- Verify with `echo %OneDrive%` or `echo %OneDriveCommercial%` in a terminal

### "Share" opens File Explorer instead of the share dialog

- The OneDrive shell integration may not be loaded. Restart the OneDrive sync client.
- Some enterprise configurations disable shell integration — contact your IT admin.

### "Open in Word" fails or opens the wrong app

- Verify the registry key exists: `HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE`
- If you have multiple Office installations, the extension uses whichever is registered in `App Paths`.
- Paths with spaces (e.g., `OneDrive - Contoso`) are properly quoted. If you still see errors, check the VS Code Developer Console (`Help` → `Toggle Developer Tools`) for details.

### "Open on Web" shows a warning about web endpoint

- The OneDrive web endpoint must be present in `HKCU:\Software\Microsoft\OneDrive\Accounts\*`
- Personal OneDrive accounts may not have a `ServiceEndpointUri` — business accounts typically do.
- The extension falls back to the shell "View online" verb if the registry entry is missing.

### Office preview shows "Preview not available"

- The OneDrive web endpoint is required for the embedded preview.
- The iframe embed requires your browser session to be signed in to your SharePoint tenant. Authentication cookies are not shared with VS Code webviews.
- Use the toolbar buttons (**Open in App**, **Open on Web**) as reliable alternatives.

---

## Known Limitations

| Limitation | Details |
|------------|---------|
| **Windows only** | macOS and Linux support is planned for a future release |
| **Web endpoint required** | "Open on Web" and Office preview need the OneDrive registry endpoint |
| **Shell integration dependency** | "Share" relies on the OneDrive sync client's shell verb |
| **Webview authentication** | Office Online iframe may not authenticate in VS Code's webview — use action buttons as fallback |
| **No inline editing** | The Office preview is read-only; edits require opening in the desktop app or web |

---

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup, architecture overview, coding guidelines, and how to submit changes.

---

## License

MIT
