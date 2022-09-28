import { ChildProcess } from 'child_process';
import * as fse from 'fs-extra';
import * as vscode from 'vscode';
import { window } from 'vscode';
import * as stBar from './statusBar';
import * as iconv from 'iconv-lite';

export class VbecmTextDocumentContentProvider implements vscode.TextDocumentContentProvider {
  // emitter and its event
  onDidChangeEmitter = new vscode.EventEmitter<vscode.Uri>();
  onDidChange = this.onDidChangeEmitter.event;

  refresh(uri: vscode.Uri): void {
    this.onDidChangeEmitter.fire(uri);
  }

  async provideTextDocumentContent(uriModule: vscode.Uri): Promise<string> {
    const isExist = await fileExists(uriModule.path);
    const text = isExist ? fse.readFileSync(uriModule.path) : undefined;
    const localeText = text ? s2u(text) : '';
    return localeText;
  }
}

export async function fileExists(filepath: string) {
  try {
    const res = (await fse.promises.lstat(filepath)).isFile();
    return res;
  } catch (e) {
    return false;
  }
}

export async function dirExists(filepath: string) {
  try {
    const res = (await fse.promises.lstat(filepath)).isDirectory();
    return res;
  } catch (e) {
    return false;
  }
}

export function myLog(message: string, title: string = '(_empty_)', isOut: boolean = true) {
  if (!isOut) {
    return;
  }
  console.log(`vbecm-ts: ${title}: ${message}`);
}

export async function rmDirIfExist(pathFolder: string, option: {}) {
  try {
    const isExist = await dirExists(pathFolder);
    if (!isExist) {
      // no folder no delete
      return;
    }
    await fse.promises.rm(pathFolder, option);
  } catch (error) {
    throw error;
  }
}

export async function rmFileIfExist(pathFile: string, option: { force: true }) {
  try {
    const isExist = await fileExists(pathFile);
    if (!isExist) {
      // no file no delete
      return;
    }
    await fse.promises.rm(pathFile, option);
  } catch (error) {
    throw error;
  }
}

export function displayMenus(isOn: boolean = true) {
  vscode.commands.executeCommand('setContext', 'vbecm.showVbsCommand', isOn);
  stBar.updateStatusBarItem(!isOn);
}

function handleSpawnResult(ls: ChildProcess): Promise<boolean> {
  return new Promise<boolean>((resolve, reject) => {
    ls.stdout?.on('data', (data) => {
      console.log(`stdout: ${s2u(data)}`);
    });

    ls.stderr?.on('data', (data) => {
      console.log(`stderr: ${s2u(data)}`);
      window.showErrorMessage(`Some error occurred: ${data}`);
      return resolve(false);
    });

    ls.on('close', (code) => {
      console.log(`child process exited with code ${code}`);
      return resolve(true);
    });
  });
}

function s2u(sb: Buffer) {
  const vbsEncode =
    vscode.workspace.getConfiguration('vbecm').get<string>('vbsEncode') || 'windows-31j';
  return iconv.decode(sb, vbsEncode);
}
