import * as path from "path";
import * as vscode from "vscode";
import { execFile } from "child_process";

/**
 * Check if the MarkMyWord CLI is available.
 */
async function isMarkMyWordInstalled(): Promise<boolean> {
  return new Promise((resolve) => {
    execFile("markmyword", ["--version"], { timeout: 10000 }, (err) => {
      resolve(!err);
    });
  });
}

/**
 * Install or update the MarkMyWord CLI as a .NET global tool via the integrated terminal.
 * Returns true if the install/update command completed successfully.
 */
async function installOrUpdateMarkMyWord(isUpdate: boolean): Promise<boolean> {
  const terminal = vscode.window.createTerminal("Paperclipped");
  terminal.show();
  const action = isUpdate ? "update" : "install";
  terminal.sendText(
    `dotnet tool ${action} -g specworks.markmyword.cli`
  );

  // Wait for the terminal to finish by polling for the CLI
  const maxWait = 60000;
  const interval = 2000;
  let elapsed = 0;
  while (elapsed < maxWait) {
    await new Promise((r) => setTimeout(r, interval));
    elapsed += interval;
    if (await isMarkMyWordInstalled()) {
      return true;
    }
  }
  return false;
}

/**
 * Convert a markdown file to Word (.docx) using the MarkMyWord CLI.
 */
async function convertToWord(
  inputPath: string,
  outputPath: string
): Promise<void> {
  return new Promise((resolve, reject) => {
    execFile(
      "markmyword",
      ["convert", "-i", inputPath, "-o", outputPath, "--force"],
      { timeout: 30000 },
      (err, _stdout, stderr) => {
        if (err) {
          reject(new Error(stderr || err.message));
        } else {
          resolve();
        }
      }
    );
  });
}

/**
 * Export the current markdown file to a Word document.
 * The .docx is saved alongside the .md file.
 */
export async function exportToWord(filePath: string): Promise<void> {
  const ext = path.extname(filePath).toLowerCase();
  if (ext !== ".md" && ext !== ".markdown") {
    vscode.window.showWarningMessage(
      "Export to Word is only available for Markdown files."
    );
    return;
  }

  // Check if MarkMyWord is installed — auto-install if needed, or update to latest
  let installed = await isMarkMyWordInstalled();
  if (!installed) {
    const install = await vscode.window.showInformationMessage(
      "MarkMyWord CLI is required for Markdown → Word conversion. Install it now?",
      "Install",
      "Cancel"
    );
    if (install !== "Install") {
      return;
    }

    const success = await vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: "Installing MarkMyWord CLI...",
        cancellable: false,
      },
      async () => {
        return installOrUpdateMarkMyWord(false);
      }
    );

    if (!success) {
      vscode.window.showErrorMessage(
        "Failed to install MarkMyWord. Make sure the .NET SDK is installed and try: dotnet tool install -g specworks.markmyword.cli"
      );
      return;
    }

    installed = await isMarkMyWordInstalled();
    if (!installed) {
      vscode.window.showErrorMessage(
        "MarkMyWord installed but not found in PATH. You may need to restart VS Code."
      );
      return;
    }
  } else {
    // Already installed — silently update to latest in the background
    execFile(
      "dotnet",
      ["tool", "update", "-g", "specworks.markmyword.cli"],
      { timeout: 60000 },
      () => { /* fire-and-forget */ }
    );
  }

  // Build output path
  const dir = path.dirname(filePath);
  const baseName = path.basename(filePath, ext);
  const outputPath = path.join(dir, `${baseName}.docx`);

  try {
    await vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: "Exporting to Word...",
        cancellable: false,
      },
      async () => {
        await convertToWord(filePath, outputPath);
      }
    );

    const action = await vscode.window.showInformationMessage(
      `Exported to ${path.basename(outputPath)}`,
      "Open in Word",
      "Reveal in Explorer"
    );

    if (action === "Open in Word") {
      const uri = vscode.Uri.file(outputPath);
      vscode.commands.executeCommand("paperclipped.openInWord", uri);
    } else if (action === "Reveal in Explorer") {
      vscode.commands.executeCommand(
        "revealFileInOS",
        vscode.Uri.file(outputPath)
      );
    }
  } catch (err: any) {
    vscode.window.showErrorMessage(
      `Export failed: ${err.message}`
    );
  }
}
