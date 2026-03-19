# Contributing to Paperclip

Thank you for your interest in contributing! This guide covers everything you need to get started — from setting up the development environment to understanding the architecture and submitting pull requests.

---

## Table of Contents

- [Prerequisites](#prerequisites)
- [Development setup](#development-setup)
- [Project structure](#project-structure)
- [Architecture overview](#architecture-overview)
  - [Activation flow](#activation-flow)
  - [Module responsibilities](#module-responsibilities)
  - [Data flow diagrams](#data-flow-diagrams)
- [Key patterns & conventions](#key-patterns--conventions)
  - [PowerShell execution](#powershell-execution)
  - [Path handling](#path-handling)
  - [Error handling](#error-handling)
  - [Context keys](#context-keys)
- [Adding a new command](#adding-a-new-command)
- [Adding a new file type](#adding-a-new-file-type)
- [Build & package](#build--package)
- [Testing](#testing)
- [Debugging](#debugging)
- [Pull request guidelines](#pull-request-guidelines)
- [Code style](#code-style)
- [Architecture decision records](#architecture-decision-records)

---

## Prerequisites

| Tool | Version | Notes |
|------|---------|-------|
| **Node.js** | 18+ | LTS recommended |
| **npm** | 9+ | Comes with Node.js |
| **VS Code** | 1.85+ | For development and testing |
| **Windows** | 10/11 | Extension is Windows-only (for now) |
| **OneDrive sync client** | Latest | For testing sharing and web URL features |
| **Microsoft Office** | Any version | For testing "Open in" commands (optional) |

---

## Development Setup

```bash
# Clone the repo
git clone <repo-url>
cd vscode-paperclip-office

# Install dependencies
npm install

# Build once
npm run compile

# Watch mode (auto-rebuild on save)
npm run watch
```

### Running in Development

1. Open the project in VS Code
2. Press **F5** to launch the **Extension Development Host**
3. In the new window, open a folder that's inside a OneDrive-synced directory
4. Right-click a file → you should see the **Paperclip** submenu

### Quick iteration cycle

```bash
# Terminal 1: Watch mode
npm run watch

# Terminal 2: Test changes
# Just reload the Extension Development Host (Ctrl+Shift+P → "Reload Window")
```

---

## Project Structure

```
vscode-paperclip-office/
├── src/
│   ├── extension.ts      # Entry point — activation, command registration
│   ├── onedrive.ts        # OneDrive detection, registry queries, URL building
│   ├── sharing.ts         # Command implementations (share, open in app, open on web)
│   └── preview.ts         # Custom editor provider for Office document preview
├── resources/
│   └── icons/
│       ├── share.svg      # Share command icon
│       ├── word.svg       # Word icon (blue W)
│       ├── excel.svg      # Excel icon (green X)
│       ├── powerpoint.svg # PowerPoint icon (red P)
│       └── web.svg        # Open on Web icon (globe)
├── dist/                  # Build output (esbuild bundle)
│   └── extension.js       # Single bundled file
├── package.json           # Extension manifest (commands, menus, editors)
├── tsconfig.json          # TypeScript configuration
├── esbuild.js             # Build script
├── .vscodeignore          # Files excluded from VSIX package
├── README.md              # User-facing documentation
└── CONTRIBUTING.md         # This file
```

---

## Architecture Overview

### Activation Flow

```
VS Code starts
  ↓
onStartupFinished event fires
  ↓
extension.ts: activate()
  ↓
├─ Check process.platform === "win32" (bail if not Windows)
├─ refreshOneDriveContext() → set context key for menu visibility
├─ Register workspace folder change listener
├─ Register OfficePreviewProvider (custom editor)
├─ Register 5 commands (share, openInWord, openInExcel, openInPowerPoint, openOnWeb)
└─ Log discovered OneDrive roots
```

### Module Responsibilities

#### `extension.ts` — Orchestrator

- Activation gate (Windows-only)
- Context key management (`paperclip:isOneDriveWorkspace`)
- Command registration and URI resolution
- Glue between VS Code API and domain modules

#### `onedrive.ts` — OneDrive Domain

- **Discovery**: Finds OneDrive roots from env vars + filesystem scan
- **Classification**: Business vs. personal accounts
- **Registry access**: Reads web endpoints from `HKCU:\Software\Microsoft\OneDrive\Accounts`
- **URL building**: Constructs SharePoint `Doc.aspx` URLs from relative file paths
- **Utilities**: PowerShell runner, string escaping

Key types:
```typescript
interface OneDriveAccount {
  localPath: string;           // e.g., "C:\Users\user\OneDrive - Microsoft"
  accountType: "business" | "personal";
  accountName: string;         // e.g., "OneDrive - Microsoft"
  webEndpoint?: string;        // e.g., "https://contoso-my.sharepoint.com/personal/..."
}
```

#### `sharing.ts` — Command Implementations

Three core functions, each with fallback strategies:

| Function | Primary strategy | Fallback |
|----------|-----------------|----------|
| `shareFile()` | Shell COM `InvokeVerb("Share")` | Open File Explorer with file selected |
| `openInOfficeApp()` | Registry `App Paths` → `Start-Process` | `Start-Process` on file (default handler) |
| `openOnWeb()` | Build SharePoint URL → `vscode.env.openExternal` | Shell COM `InvokeVerb("View online")` |

#### `preview.ts` — Custom Editor

- Implements `CustomReadonlyEditorProvider`
- Maps file extensions to app metadata (name, color, icon letter)
- Builds HTML with toolbar + iframe
- Handles webview ↔ extension message passing

### Data Flow Diagrams

#### Share flow
```
User right-clicks file → share command → shareFile(path)
  → psEscape(folder, file)
  → runPowerShell(Shell.Application script)
  → COM finds "Share" verb → DoIt()
  → OneDrive native share dialog appears
```

#### Open in Office flow
```
User right-clicks .docx → openInWord command → openInOfficeApp(path, "word")
  → Registry lookup: HKLM\App Paths\WINWORD.EXE
  → Start-Process -FilePath <app> -ArgumentList '"<file>"'
  → Word desktop opens the file
```

#### Web URL flow
```
User right-clicks file → openOnWeb command → openOnWeb(path)
  → findOneDriveRoot(path) → match account
  → enrichAccountsFromRegistry() → read ServiceEndpointUri
  → buildWebUrl(path, account) → Doc.aspx URL
  → vscode.env.openExternal(url) → browser opens
```

---

## Key Patterns & Conventions

### PowerShell Execution

All Windows system interactions go through `runPowerShell()` in `onedrive.ts`:

```typescript
runPowerShell(script: string): Promise<string>
```

- Uses `powershell.exe` with `-NoProfile -NonInteractive -ExecutionPolicy Bypass`
- 15-second timeout
- Returns stdout; throws on non-zero exit or stderr

**When writing PowerShell scripts:**
- Use `Write-Output` for structured return values (e.g., `'OK'`, `'NO_SHARE_VERB'`)
- Check return codes in TypeScript, not in PowerShell
- Escape file paths with `psEscape()` (doubles single quotes: `'` → `''`)

### Path Handling

**Critical**: Paths containing spaces (e.g., `OneDrive - Microsoft`) require special quoting in PowerShell:

```typescript
// ✅ Correct: double-quote wrapping for Start-Process arguments
Start-Process -FilePath $appPath -ArgumentList ('"' + $filePath + '"')

// ❌ Wrong: unquoted path will split on spaces
Start-Process -FilePath $appPath -ArgumentList $filePath
```

Always use `psEscape()` for embedding paths in single-quoted strings and explicit double-quote wrapping for `Start-Process -ArgumentList`.

### Error Handling

- **User-facing errors**: Use `vscode.window.showErrorMessage()` / `showWarningMessage()` / `showInformationMessage()`
- **Silent failures**: Log to console, don't interrupt the user
- **Fallback pattern**: Every command has a fallback strategy (see sharing.ts)

### Context Keys

The extension uses a single VS Code context key:

```
paperclip:isOneDriveWorkspace = true | false
```

This controls the visibility of the `Paperclip` submenu in the explorer context menu. It's refreshed on activation and whenever workspace folders change.

---

## Adding a New Command

1. **Define the command** in `package.json` under `contributes.commands`:
   ```json
   {
     "command": "paperclip.myCommand",
     "title": "My Command",
     "category": "Paperclip",
     "icon": {
       "light": "resources/icons/my-icon.svg",
       "dark": "resources/icons/my-icon.svg"
     }
   }
   ```

2. **Add menu entry** in `package.json` under `contributes.menus["paperclip.menu"]`:
   ```json
   {
     "command": "paperclip.myCommand",
     "when": "<optional context expression>",
     "group": "2_openApp"
   }
   ```
   Groups: `1_share`, `2_openApp`, `3_web` (controls separator placement)

3. **Implement the command** in `sharing.ts` (or a new module if the scope warrants it):
   ```typescript
   export async function myCommand(filePath: string): Promise<void> {
     // Implementation
   }
   ```

4. **Register in `extension.ts`**:
   ```typescript
   vscode.commands.registerCommand(
     "paperclip.myCommand",
     async (uri?: vscode.Uri) => {
       const filePath = resolveFilePath(uri);
       if (filePath) {
         await myCommand(filePath);
       }
     }
   );
   ```

5. **Add an SVG icon** to `resources/icons/` if needed (16×16 viewBox, single path preferred)

---

## Adding a New File Type

To add support for a new Office file type in the preview and context menus:

1. **Add the `when` clause** to the relevant menu entry in `package.json`:
   ```json
   "when": "resourceExtname == .myext || resourceExtname == .docx"
   ```

2. **Add to `customEditors` selector** in `package.json`:
   ```json
   { "filenamePattern": "*.myext" }
   ```

3. **Add to `OFFICE_MAP`** in `preview.ts`:
   ```typescript
   ".myext": { name: "MyApp", color: "#hexcolor", letter: "M", commandId: "paperclip.openInMyApp", webAction: "view" },
   ```

---

## Build & Package

```bash
# Development build
npm run compile

# Watch mode (rebuilds on file changes)
npm run watch

# Production build (minified)
npm run vscode:prepublish

# Package as VSIX
npx vsce package --allow-missing-repository

# Install locally
code --install-extension vscode-paperclip-office-*.vsix --force
```

### Build system

The extension uses **esbuild** for fast bundling:

- Entry: `src/extension.ts`
- Output: `dist/extension.js` (single file, CommonJS)
- External: `vscode` (provided by the VS Code runtime)
- Source maps included in development builds

---

## Testing

### Manual testing checklist

Before submitting a PR, verify these scenarios:

- [ ] **OneDrive detection**: Open a OneDrive folder — submenu appears
- [ ] **Non-OneDrive folder**: Open a regular folder — submenu does NOT appear
- [ ] **Share**: Right-click → Paperclip → Share… → native dialog opens
- [ ] **Open in Word**: Right-click `.docx` → Open in Word → Word desktop opens
- [ ] **Open in Excel**: Right-click `.xlsx` → Open in Excel → Excel desktop opens
- [ ] **Open in PowerPoint**: Right-click `.pptx` → Open in PowerPoint → PowerPoint desktop opens
- [ ] **Open on Web**: Right-click any file → Open on Web → browser opens SharePoint URL
- [ ] **Office preview**: Double-click `.docx` → preview opens with toolbar
- [ ] **Path with spaces**: Test with files in `OneDrive - Microsoft` (spaces in path)
- [ ] **Fallback behaviors**: Test with OneDrive stopped, Office uninstalled, etc.

### Automated testing

> **TODO**: The extension does not currently have automated tests. Contributions to add a test suite using `@vscode/test-electron` are welcome.

Recommended test structure:

```
test/
├── suite/
│   ├── onedrive.test.ts     # Unit tests for detection and URL building
│   ├── sharing.test.ts      # Unit tests for command logic
│   └── extension.test.ts    # Integration tests
└── runTest.ts               # Test runner
```

---

## Debugging

### VS Code Developer Tools

1. In the Extension Development Host: **Help** → **Toggle Developer Tools**
2. Check the **Console** tab for `[Paperclip]` messages
3. PowerShell errors and fallback triggers are logged here

### Breakpoints

1. Set breakpoints in `src/*.ts` files
2. Press **F5** — the debugger attaches to the Extension Development Host
3. Trigger a command and step through the code

### PowerShell script debugging

To debug a PowerShell script independently:

1. Copy the script from `sharing.ts` or `onedrive.ts`
2. Replace `psEscape()` substitutions with literal values
3. Run in a PowerShell terminal to see raw output

---

## Pull Request Guidelines

### Before submitting

- [ ] Code compiles without errors (`npm run compile`)
- [ ] `package.json` is valid JSON (easy to break with comma errors)
- [ ] All manual test scenarios pass
- [ ] No hardcoded user-specific paths or credentials
- [ ] README and CONTRIBUTING updated if new features were added

### PR structure

- **Title**: Short, imperative (e.g., "Add copy link command", "Fix path quoting on share")
- **Description**: What changed, why, and how to test
- **One concern per PR**: Don't mix features with refactors

### Commit messages

```
<type>: <short description>

<optional body>

Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>
```

Types: `feat`, `fix`, `refactor`, `docs`, `chore`, `test`

---

## Code Style

- **TypeScript** strict mode
- **No semicolons** omission — always use semicolons
- **Double quotes** for strings (consistent with existing code)
- **Async/await** over raw Promises
- **Explicit return types** on exported functions
- **No `any`** unless absolutely necessary (use `unknown` + type narrowing)
- **Named exports** only (no default exports)

---

## Architecture Decision Records

### ADR-001: PowerShell for system interactions

**Decision**: Use `child_process.execFile("powershell.exe")` for all Windows system interactions (COM objects, registry, process launching).

**Rationale**: Node.js doesn't have native COM support. The `edge-js` or `winax` packages require native compilation and would complicate the build. PowerShell provides direct access to COM, WMI, and the registry with no dependencies.

**Tradeoff**: ~100–200ms overhead per PowerShell invocation (process startup). Acceptable for user-initiated actions.

### ADR-002: Shell.Application COM for sharing

**Decision**: Use `Shell.Application.NameSpace().ParseName().Verbs()` to find the OneDrive share verb, rather than calling a OneDrive API.

**Rationale**: OneDrive doesn't expose a public sharing API for local files. The shell verb is the same mechanism File Explorer uses and supports all OneDrive configurations (personal, business, multi-account).

### ADR-003: Registry-based Office app resolution

**Decision**: Read Office executable paths from `HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\` rather than hardcoding paths or using `where.exe`.

**Rationale**: `App Paths` is the canonical Windows mechanism for resolving application paths. It works across Office versions (2016, 2019, 365) and installation types (MSI, Click-to-Run).

### ADR-004: esbuild over webpack

**Decision**: Use esbuild for bundling instead of webpack.

**Rationale**: Sub-second builds, simpler configuration, and sufficient for a single-entry-point extension. No need for webpack's plugin ecosystem.

### ADR-005: Custom editor for Office preview

**Decision**: Implement `CustomReadonlyEditorProvider` with an Office Online iframe, with action buttons as fallback.

**Rationale**: VS Code webviews don't share browser authentication cookies, so the iframe may not work for all users. The action buttons (Open in App, Open on Web, Share) provide a reliable alternative. Future improvement: use `mammoth.js` for local `.docx` → HTML rendering.
