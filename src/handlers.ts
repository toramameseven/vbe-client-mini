import * as path from 'path';
import * as vscode from 'vscode';
import { TextEditor } from 'vscode';
import * as stBar from './statusBar';
import * as vbs from './vbsModule';
import * as vbecmCommon from './vbecmCommon';
import { DiffFileInfo, FileDiff } from './diffFiles';
import { showInformationMessage, showErrorMessage, showWarningMessage } from './vbecmCommon';

const STRING_EMPTY = '';

const testAndConfirm =
  (message: string) =>
  async (baseDir: string, srcDir: string, targetDir: string, compareTo: 'src' | 'vbe') => {
    const ans = await vbs.comparePathAndGo(baseDir, srcDir, targetDir, compareTo, message);
    return ans;
  };

// for book
// export all modules
export async function handlerExportModulesFromBook(uriBook: vscode.Uri) {
  // function to test source is modified?
  const diffTestAndConfirm = testAndConfirm(
    'Some files modified in the Source folder. Do you export? Check vbe view.'
  );

  displayMenus(false);
  // export the source folder and the base folder
  try {
    const r = await vbs.exportModulesAndSynchronize(
      uriBook.fsPath,
      STRING_EMPTY,
      diffTestAndConfirm
    );
    r && showInformationMessage('Success export modules.');

    await vbs.updateModification(uriBook.fsPath);
  } catch (e) {
    showErrorMessage('Error export modules.');
    showErrorMessage(e);
  } finally {
    displayMenus();
  }
}

export function handlerCompile(uriBook: vscode.Uri) {
  displayMenus(false);
  try {
    vbs.compile(uriBook.fsPath);
    showInformationMessage('Success compile, may be.');
  } catch (e) {
    showInformationMessage('Error compile.');
    showErrorMessage(e);
  } finally {
    displayMenus(true);
  }
}
// import all modules
export async function handlerImportModulesToBook(uriBook: vscode.Uri) {
  // temp vs vase
  const diffTestAndConfirm = testAndConfirm(
    'Excel Vba may be modified. Do you import? Check vbe view.'
  );

  displayMenus(false);
  try {
    const r = await vbs.importModules(uriBook.fsPath, STRING_EMPTY, diffTestAndConfirm);
    r && showInformationMessage('Success import modules.');

    await vbs.updateModification();
  } catch (e) {
    showErrorMessage('Error import modules.');
    showErrorMessage(e);
  } finally {
    displayMenus(true);
  }
}

export async function handlerExportFrxModulesFromBook(uriBook: vscode.Uri) {
  displayMenus(false);
  try {
    await vbs.exportFrxModules(uriBook.fsPath);
    showInformationMessage('Success export frx modules.');
  } catch (e) {
    showErrorMessage('Error export frx modules.');
    showErrorMessage(e);
  } finally {
    displayMenus();
  }
}

// for module
// export a module
export async function handlerPullModuleAsync(uriModule: vscode.Uri) {
  // if the module file does not exist, return STRING_EMPTY
  const bookPath = await getExcelPathFromModule(uriModule);
  if (bookPath === undefined) {
    showWarningMessage(uriModule.fsPath + ' is not a VBE project file.');
    return;
  }

  // todo check local modified
  const diffTestAndConfirm = testAndConfirm(
    'This file is modified. Do you want to pull a file from excel?'
  );

  displayMenus(false);
  try {
    const modulePath = uriModule.fsPath;
    const r = await vbs.exportModulesAndSynchronize(bookPath, modulePath, diffTestAndConfirm);
    r && showInformationMessage('Success pull.');

    await vbs.updateModification();
  } catch (e) {
    showErrorMessage('Error pull.');
    showErrorMessage(e);
  } finally {
    displayMenus(true);
  }
}

// commit a module
export async function handlerCommitModuleFromFile(uriModule: vscode.Uri) {
  const diffTestAndConfirm = testAndConfirm('Excel Vba may be modified. Do you push this file?');

  const bookPath = await getExcelPathFromModule(uriModule);
  if (bookPath === undefined) {
    showWarningMessage(uriModule.fsPath + ' is not a VBE project file.');
    return;
  }

  displayMenus(false);
  try {
    const modulePath = uriModule.fsPath;
    const r = await vbs.importModules(bookPath, modulePath, diffTestAndConfirm);
    r && showInformationMessage('Success push.');

    await vbs.updateModification();
  } catch (e) {
    showErrorMessage('Error push.');
    showErrorMessage(e);
  } finally {
    displayMenus(true);
  }
}

export async function handlerVbaRunFromEditor(
  textEditor: TextEditor,
  edit: vscode.TextEditorEdit,
  uriModule: vscode.Uri
) {
  // excel file path
  const bookPath = await getExcelPathFromModule(uriModule);
  if (bookPath === undefined) {
    showWarningMessage(uriModule.fsPath + ' is not a VBE project file.');
    return;
  }

  // get sub function
  const activeLine = textEditor.document.lineAt(textEditor.selection.active.line).text;
  const vbaSub = activeLine.match(/Sub (.*)\(\s*\)/i);
  const funcName = vbaSub !== null && vbaSub.length === 2 && vbaSub[1];
  if (funcName === false) {
    showErrorMessage('Error run. No sub name.');
    return;
  }

  displayMenus(false);
  try {
    const modulePath = uriModule.fsPath;
    await vbs.vbaSubRun(bookPath, path.basename(modulePath), funcName);
    showInformationMessage('Success run, may be.');
  } catch (e) {
    showErrorMessage('Error run.');
    showErrorMessage(e);
  } finally {
    displayMenus(true);
  }
}

// commit all modules for commit
export async function handlerCommitAllModuleFromFolder(uriFolder: vscode.Uri) {
  const bookPath = await getExcelPathSrcFolder(uriFolder);
  if (bookPath === undefined) {
    showWarningMessage(uriFolder.fsPath + ' is not a VBE project folder.');
    return;
  }

  const diffTestAndConfirm = testAndConfirm(
    'Excel Vba may be modified. Do you push all? Check vbe view.'
  );

  displayMenus(false);
  try {
    const r = await vbs.importModules(bookPath, STRING_EMPTY, diffTestAndConfirm);
    r && showInformationMessage('Success push all.');
    await vbs.updateModification();
  } catch (e) {
    showErrorMessage('Error push all.');
    showErrorMessage(e);
  } finally {
    displayMenus(true);
  }
}

export async function handlerCheckModifiedOnFolder(uriFolder: vscode.Uri) {
  displayMenus(false);
  try {
    const bookPath = await getExcelPathSrcFolder(uriFolder);
    await vbs.updateModification(bookPath);
  } catch (e) {
    showErrorMessage('Error CheckModified.');
    showErrorMessage(e);
  } finally {
    displayMenus();
  }
}

// vbe diff tree view
export async function handlerUpdateModificationOnFolder() {
  displayMenus(false);
  try {
    await vbs.updateModification();
  } catch (error) {
    showErrorMessage('error update modification');
    showErrorMessage(error);
  } finally {
    displayMenus();
  }
}

export async function handlerCommitModuleFromVbeDiff(fileInfo: DiffFileInfo) {
  const bookPath = await getExcelPathFromModule(vscode.Uri.file(fileInfo.compareFilePath!));
  if (bookPath === undefined) {
    showWarningMessage(fileInfo.moduleName + ' is not a VBE project file.');
    return;
  }

  const isRemove = !(await vbecmCommon.fileExists(fileInfo.compareFilePath!));

  if (isRemove) {
    await handlerCommitModuleFromVbeDiffRemove(bookPath, fileInfo.moduleName!);
  } else {
    await handlerCommitModuleFromFile(vscode.Uri.file(fileInfo.compareFilePath!));
  }
}

async function handlerCommitModuleFromVbeDiffRemove(bookPath: string, modulePath: string) {
  try {
    await vbs.removeModuleSync(bookPath, modulePath);
    await vbs.updateModification(bookPath);
    showInformationMessage('Success push(remove module).');
  } catch (error) {
    showErrorMessage('error push(remove module)');
    showErrorMessage(error);
  } finally {
    displayMenus();
  }
  return true;
}

export async function handlerResolveVbeConflicting(fileInfo: DiffFileInfo) {
  const bookPath = await getExcelPathFromVbeModule(vscode.Uri.file(fileInfo.compareFilePath!));
  if (bookPath === undefined) {
    showWarningMessage(fileInfo.compareFilePath + ' is not a VBE project file.');
    return;
  }

  displayMenus(false);
  try {
    await vbs.resolveVbeConflicting(bookPath, fileInfo.moduleName!);
  } catch (error) {
    showErrorMessage('Error CheckModified.');
    showErrorMessage(error);
  } finally {
    displayMenus();
  }
}

export async function handlerDiffBaseTo(resource: DiffFileInfo): Promise<void> {
  await vbs.diffBaseTo(resource);
}

export async function handlerDiffSrcToVbe(resource: DiffFileInfo): Promise<void> {
  const pathBook = await getExcelPathFromVbeModule(vscode.Uri.file(resource.compareFilePath!));
  if (pathBook === undefined) {
    showWarningMessage(resource.compareFilePath + ' is not a VBE project file.');
    return;
  }
  await vbs.diffSrcToVbe(resource, pathBook);
}

export async function handlerCollapseAllVbeDiffView() {
  await vscode.commands.executeCommand('workbench.actions.treeView.vbeDiffView.collapseAll');
}
// common
export function displayMenus(isOn: boolean = true) {
  vscode.commands.executeCommand('setContext', 'vbecm.showVbsCommand', isOn);
  stBar.updateStatusBarItem(!isOn);
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
  const isFile = await vbecmCommon.fileExists(xlsPath);
  return isFile ? xlsPath : undefined;
}
