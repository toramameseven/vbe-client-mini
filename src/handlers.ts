import * as path from 'path';
import * as vscode from 'vscode';
import { TextEditor } from 'vscode';
import * as stBar from './statusBar';
import * as vbs from './vbsModule';
import * as common from './common';
import { rawListeners } from 'process';
import { DiffFileInfo, FileDiff } from './diffFiles';

const STRING_EMPTY = '';


const testAndConfirm = 
  (message : string) => 
  async (baseDir: string, srcDir: string, targetDir: string, compareTo : 'src' |'vbe') => {
    const ans = await vbs.comparePathAndGo(baseDir, srcDir, targetDir, compareTo, message);
    return ans;
  };

export async function handlerExportModules(uriBook: vscode.Uri){
  displayMenus(false);

  // function to test source is modified?
  const diffTestAndConfirm = testAndConfirm('Some files modified in the Source folder. Do you export? Check vbe view.');
  
  // export the source folder and the base folder
  try{
    const r = await vbs.exportModulesForUpdate(uriBook.fsPath, STRING_EMPTY, diffTestAndConfirm);
    r && showInformationMessage('Success export modules.');

    await vbs.updateModification();
  }
  catch(e)
  {
    showErrorMessage('Error export modules.');
    showErrorMessage(e);
  }
  finally {
    displayMenus(true);
  }
};


export async function handlerUpdateAsync(uriModule: vscode.Uri) {
  displayMenus(false);

  // if the module file does not exist, return STRING_EMPTY
  const bookPath = await getExcelPathFromModule(uriModule);
  const modulePath = uriModule.fsPath;
  if (bookPath === STRING_EMPTY) {
    showErrorMessage(`Excel file does not exist to update.: ${modulePath}`);
    displayMenus(true);
    return;
  }

  // todo check local modified
  const diffTestAndConfirm = testAndConfirm('This file is modified. Do you want to pull a file from excel?');

  try {
    const r = await vbs.exportModulesForUpdate(bookPath, modulePath, diffTestAndConfirm);
    r && showInformationMessage('Success pull.');

    await vbs.updateModification();
  } catch (e) {
    showErrorMessage('Error pull.');
    showErrorMessage(e);
  }
  finally {
    displayMenus(true);
  }
}


export async function handlerImportModules(uriBook: vscode.Uri) {
  displayMenus(false);
  
  // temp vs vase
  const diffTestAndConfirm = testAndConfirm('Excel Vba may be modified. Do you import? Check vbe view.');

  try {
    const r = await vbs.importModules(uriBook.fsPath, STRING_EMPTY, diffTestAndConfirm);
    r &&  showInformationMessage('Success import modules.');

    await vbs.updateModification();

  } catch (e) {
    showErrorMessage('Error import modules.');
    showErrorMessage(e);
  }
  finally {
    displayMenus(true);
  }
}


export async function handlerCommitAllModule(uriFolder: vscode.Uri) {
  const bookPath = await getExcelPathSrcFolder(uriFolder);
  displayMenus(false);

  const diffTestAndConfirm = testAndConfirm('Excel Vba may be modified. Do you push all? Check vbe view.');

  try {
    const r = await vbs.importModules(bookPath, STRING_EMPTY, diffTestAndConfirm);
    r && showInformationMessage('Success push all.');

    await vbs.updateModification();
  } catch (e) {
    showErrorMessage('Error push all.');
    showErrorMessage(e);
  }
  finally {
    displayMenus(true);
  }
}

export async function handlerCommitModule(uriModule: vscode.Uri) {
  displayMenus(false);
  const diffTestAndConfirm = testAndConfirm('Excel Vba may be modified. Do you push this file?');

  const bookPath = await getExcelPathFromModule(uriModule);
  const modulePath = uriModule.fsPath;
  if (await common.dirExists(bookPath)) {
    showErrorMessage(`Excel file does not exist to commit.: ${modulePath}`);
    displayMenus(true);
    return;
  }

  try {
    const r = await vbs.importModules(bookPath, modulePath, diffTestAndConfirm);
    r && showInformationMessage('Success push.');

    await vbs.updateModification();
  } catch (e) {
    showErrorMessage('Error push.');
    showErrorMessage(e);
  }
  finally {
    displayMenus(true);
  }
}

export async function handlerCommitModuleFromVbeDiff(fileInfo: DiffFileInfo) {
  await handlerCommitModule(vscode.Uri.file(fileInfo.compareFilePath!));
}


export async function handlerUpdateFrxModules(uriBook:vscode.Uri){
  displayMenus(false);
  try {
    await vbs.exportFrxModules(uriBook.fsPath);
    showInformationMessage('Success export frx modules.');
  } catch (e) {
    showErrorMessage('Error export frx modules.');
    showErrorMessage(e);
  }
  displayMenus();
}

export function handlerCompile(uriBook: vscode.Uri) {
  displayMenus(false);
  try {
    vbs.compile(uriBook.fsPath);
    showInformationMessage('Success compile, may be.');
  } catch (e) {
    showInformationMessage('Error compile.');
    showErrorMessage(e);
  }
  finally {
    displayMenus(true);
  }
}


export async function handlerVbaRun(textEditor: TextEditor, edit: vscode.TextEditorEdit, uriModule: vscode.Uri) {
  displayMenus(false);
  // excel file path
  const bookPath = await getExcelPathFromModule(uriModule);
  const modulePath = uriModule.fsPath;

  // get sub function
  const activeLine = textEditor.document.lineAt(textEditor.selection.active.line).text;
  const vbaSub = activeLine.match(/Sub (.*)\(\s*\)/i);
  const funcName = vbaSub !== null && vbaSub.length === 2 && vbaSub[1];

  if (funcName === false){
    showErrorMessage('Error run. No sub name.');
    displayMenus(true);
    return;
  }

  try {
    await vbs.vbaSubRun(bookPath, path.basename(modulePath), funcName);
    showInformationMessage('Success run, may be.');
  } catch (e) {
    showErrorMessage('Error run.');
    showErrorMessage(e);
  }
  finally {
    displayMenus(true);
  }
}

export async function handlerCheckModified(uriFolder: vscode.Uri)
{
  const bookPath = await getExcelPathSrcFolder(uriFolder);
  await vbs.updateModification(bookPath);
}

export async function handlerResolveVbeConflicting(fileInfo: DiffFileInfo){
  // get book path from a modules in vbe folder.
  const bookPath = await getExcelPathFromVbeModule(vscode.Uri.file(fileInfo.compareFilePath!));
  await vbs.resolveVbeConflicting(bookPath, fileInfo.moduleName!);
}

export function displayMenus(isOn: boolean = true) {
  vscode.commands.executeCommand('setContext', 'vbecm.showVbsCommand', isOn);
  stBar.updateStatusBarItem(!isOn);
}

export async function collapseAllVbeDiffView(){
  await vscode.commands.executeCommand('workbench.actions.treeView.vbeDiffView.collapseAll');
}

export async function handlerDiffBaseTo(resource: DiffFileInfo): Promise<void>{
  await vbs.diffBaseTo(resource);
}

export async function handlerDiffSrcToVbe(resource: DiffFileInfo): Promise<void>{
  const pathBook = await getExcelPathFromVbeModule(vscode.Uri.file(resource.compareFilePath!));
  await vbs.diffSrcToVbe(resource, pathBook);
}

function showInformationMessage(message: string) {
  vscode.window.showInformationMessage(getHourMinute() + ' , ' + message);
}

function showErrorMessage(message: string | unknown) {
  if (message instanceof  Error){
    vscode.window.showErrorMessage(getHourMinute() + ' , ' +  message.message);
    return;
  }
  vscode.window.showErrorMessage(getHourMinute() + ' , ' + message);
}

function getHourMinute(){
  const date = new Date();
  const dateString = date.getHours().toString().padStart(2, '0') + ':' + date.getMinutes().toString().padStart(2, '0');
  return dateString;
}

async function getExcelPathFromModule(uri: vscode.Uri) {
  const dirParent = path.dirname(uri.fsPath);
  const r = await getExcelPathSrcFolder(vscode.Uri.file(dirParent));
  return r;
}

async function getExcelPathFromVbeModule(uri: vscode.Uri) {
  const dirParent = path.dirname(path.dirname(uri.fsPath));
  const r = await getExcelPathSrcFolder(vscode.Uri.file(dirParent));
  return r;
}

async function getExcelPathSrcFolder(uriSrcFolder: vscode.Uri) {
  const dirForBook = path.dirname(uriSrcFolder.fsPath);
  const bookFileName = path.basename(uriSrcFolder.fsPath).slice('src_'.length);
  const xlsPath = path.resolve(dirForBook, bookFileName);
  const isFile = await common.fileExists(xlsPath);
  return isFile ? xlsPath : STRING_EMPTY;
}