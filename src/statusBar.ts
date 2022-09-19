import * as vscode from 'vscode';

export const statusBarVba = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Right, 100);

export function updateStatusBarItem(isVbaWork: boolean): void {
  if (isVbaWork) {
    statusBarVba.text = `$(sync~spin) vbecm`;
    statusBarVba.show();
    return;
  }
  statusBarVba.hide();
}
