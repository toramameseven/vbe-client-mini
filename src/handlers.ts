import * as path from 'path';
import * as vscode from 'vscode';
import { TextEditor } from 'vscode';
import * as stBar from './statusBar';
import * as vbs from './vbsModule';
import * as common from './common';


/**
 * Export modules
 * @param uriBook 
 */
export async function handlerExportModules(uriBook:vscode.Uri){
  displayMenus(false);

  const diffTestAndConfirm: vbs.TestConfirm = async (baseDire,srcDir, diffTitle) => {
    const ans = await vbs.comparePathAndGo(
      baseDire,
      srcDir, 
      diffTitle, 
      'Some files modified in the Source folder. Check Output Window. Do you want to export?');
    return ans === 'Yes' ? true: false;
  };

  try{
    const r = await vbs.exportModulesToScrAndBase(uriBook.fsPath, diffTestAndConfirm);
    if (r){
      showInformationMessage('Success export modules.');
    }
  }
  catch(e)
  {
    showErrorMessage('Error export modules.');
    showErrorMessage(e);
  }
  finally
  {
    displayMenus(true);
  }
};

/**
 * 
 * @param uri module path
 * @returns 
 */
 export async function handlerUpdateAsync(uriModule: vscode.Uri) {
  displayMenus(false);

  // if the file does not exist, return ''
  const xlsmPath = await getExcelPathFromModule(uriModule);
  const modulePath = uriModule.fsPath;
  if (xlsmPath === '') {
    showErrorMessage(`Excel file does not exist to checkout.: ${modulePath}`);
    displayMenus(true);
    return;
  }

  // todo check local modified
  const diffTestAndConfirm: vbs.TestConfirm = async (baseFile, srcFile, diffTitle) => {
    const ans = await vbs.comparePathAndGo(
      baseFile,
      srcFile, 
      diffTitle, 
      'file modified in the Source folder. Check Output Window. Do you want to update the file?');
    return ans === 'Yes' ? true: false;
  };

  try {
    const r = await vbs.updateModule(xlsmPath, modulePath, diffTestAndConfirm);
    r && showInformationMessage('Success update.');
  } catch (e) {
    showErrorMessage('Error update.');
    showErrorMessage(e);
  }
  finally
  {
    displayMenus(true);
  }
}

/**
 * Import modules
 * @param uriBook 
 */
export async function handlerImportModules(uriBook: vscode.Uri) {
  displayMenus(false);
  
  // temp vs vase
  const diffTestAndConfirm: vbs.TestConfirm = async (baseDir: string, targetDir: string) =>
  {
    const ans = await vbs.comparePathAndGo(
      baseDir, 
      targetDir, 
      'base, vbe', 
      'Excel Vba may be modified. Check output tab. Do you import?');
    return ans === 'Yes' ? true: false;
  };

  try {
    const r = await vbs.importModules(uriBook.fsPath, diffTestAndConfirm);
    if (r){
      showInformationMessage('Success import modules.');
    }

  } catch (e) {
    showErrorMessage('Error import modules.');
    showErrorMessage(e);
  }
  finally
  {
    displayMenus(true);
  }
}

/**
 * Commit all module
 * @param uriFolder 
 */
 export async function handlerCommitAllModule(uriFolder: vscode.Uri) {
  const xlsmPath = await getExcelPathSrcFolder(uriFolder);
  displayMenus(false);

  const diffTestAndConfirm: vbs.TestConfirm = async (baseDir: string, vbeDir: string) =>
  {
    const ans = await vbs.comparePathAndGo(
      baseDir, 
      vbeDir, 
      'base, vbe',
      'Excel Vba may be modified. Check output tab. Do you commit?');
    return ans === 'Yes' ? true: false;
  };

  try {
    const r = await vbs.importModules(xlsmPath, diffTestAndConfirm);
    r && showInformationMessage('Success commit all.');
  } catch (e) {
    showErrorMessage('Error commit all.');
    showErrorMessage(e);
  }
  finally
  {
    displayMenus(true);
  }
}

/**
 * Commit module
 * @param uriModule 
 * @returns 
 */
 export async function handlerCommitModule(uriModule: vscode.Uri) {
  displayMenus(false);

  const diffTestAndConfirm: vbs.TestConfirm = async (baseFile:string, vbeFile:string) =>
  {
    const ans = await vbs.comparePathAndGo(
      baseFile, 
      vbeFile, 
      'base, vbe',
      'Excel Vba may be modified. Check output tab. Do you commit?');
    return ans === 'Yes' ? true: false;
  };


  const xlsmPath = await getExcelPathFromModule(uriModule);
  const modulePath = uriModule.fsPath;
  if (await common.dirExists(xlsmPath)) {
    showErrorMessage(`Excel file does not exist to commit.: ${modulePath}`);
    displayMenus(true);
    return;
  }

  try {
    const r = await vbs.commitModule(xlsmPath, modulePath, diffTestAndConfirm);
    r && showInformationMessage('Success commit.');
  } catch (e) {
    showErrorMessage('Error commit.');
    showErrorMessage(e);
  }
  finally
  {
    displayMenus(true);
  }
}


/**
 * Export FRX modules
 * @param uriBook
 */
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

/**
 * Compile VBA
 * @param uriBook
 */
export function handlerCompile(uriBook: vscode.Uri) {
  displayMenus(false);
  try {
    vbs.compile(uriBook.fsPath);
    showInformationMessage('Success compile, may be.');
  } catch (e) {
    showInformationMessage('Error compile.');
    showErrorMessage(e);
  }
  finally
  {
    displayMenus(true);
  }
}


/**
 * Run sub function
 * @param textEditor
 * @param edit 
 * @param uriModule 
 * @returns 
 */
export async function handlerVbaRun(textEditor: TextEditor, edit: vscode.TextEditorEdit, uriModule: vscode.Uri) {
  displayMenus(false);
  // excel file path
  const xlsmPath = await getExcelPathFromModule(uriModule);
  const modulePath = uriModule.fsPath;

  // sub function
  const activeLine = textEditor.document.lineAt(textEditor.selection.active.line).text;
  const vbaSub = activeLine.match(/Sub (.*)\(\s*\)/i);
  const funcName = vbaSub !== null && vbaSub.length === 2 && vbaSub[1];

  if (funcName === false){
    showErrorMessage('Error run. No sub name.');
    displayMenus(true);
    return;
  }

  try {
    await vbs.vbaSubRun(xlsmPath, path.basename(modulePath), funcName);
    showInformationMessage('Success run, may be.');
  } catch (e) {
    showErrorMessage('Error run.');
    showErrorMessage(e);
  }
  finally
  {
    displayMenus(true);
  }
}










/**
 * set display vbe menu on or off
 * @param isOn
 */
export function displayMenus(isOn: boolean = true) {
  vscode.commands.executeCommand('setContext', 'vbecm.showVbsCommand', isOn);
  stBar.updateStatusBarItem(!isOn);
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

/**
 * get Excel path from a module path
 * @param uri 
 * @returns 
 */
async function getExcelPathFromModule(uri: vscode.Uri) {
  const dirParent = path.dirname(uri.fsPath);
  const dirForBook = path.dirname(dirParent);
  const bookFileName = path.basename(path.dirname(uri.fsPath)).slice('src_'.length);
  const xlsPath = path.resolve(dirForBook, bookFileName);
  const isFile = await common.fileExists(xlsPath);
  return isFile ? xlsPath : '';
}

/**
 * get Excel path from a source path
 * @param uriSrcFolder 
 * @returns 
 */
async function getExcelPathSrcFolder(uriSrcFolder: vscode.Uri) {
  const dirForBook = path.dirname(uriSrcFolder.fsPath);
  const bookFileName = path.basename(uriSrcFolder.fsPath).slice('src_'.length);
  const xlsPath = path.resolve(dirForBook, bookFileName);
  const isFile = await common.fileExists(xlsPath);
  return isFile ? xlsPath : '';
}