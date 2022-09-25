import * as vscode from 'vscode';
export const vbeOutput = vscode.window.createOutputChannel('vbecm');

// eslint-disable-next-line @typescript-eslint/naming-convention
const MessageType = {
  // eslint-disable-next-line @typescript-eslint/naming-convention
  Error: 'error',
  // eslint-disable-next-line @typescript-eslint/naming-convention
  Warning: 'warning',
  // eslint-disable-next-line @typescript-eslint/naming-convention
  Info: 'info',
} as const;
type MessageType = typeof MessageType[keyof typeof MessageType];

function createMessage(message: string | unknown) {
  let messageString;
  if (message instanceof Error) {
    return message.message;
  } else if (typeof message === 'string') {
    return message;
  }
  return 'create message error.';
}

export function showInfo(message: unknown, showNotification = true) {
  const messageOut = `[Info  - ${new Date().toLocaleTimeString()}] ${createMessage(message)}`;
  vbeOutput.appendLine(messageOut);
  if (showNotification) {
    vscode.window.showInformationMessage(messageOut);
  }
}

export function showWarn(message: unknown, showNotification = true) {
  const messageOut = `[Warn  - ${new Date().toLocaleTimeString()}] ${createMessage(message)}`;
  vbeOutput.appendLine(messageOut);
  if (showNotification) {
    vscode.window.showWarningMessage(messageOut);
  }
}

export function showError(message: unknown, showNotification = true) {
  const messageOut = `[Error  - ${new Date().toLocaleTimeString()}] ${createMessage(message)}`;
  vbeOutput.appendLine(messageOut);
  if (showNotification) {
    vscode.window.showErrorMessage(messageOut);
  }
}
