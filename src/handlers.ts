import * as path from 'path';
import * as vscode from 'vscode';
import { TextEditor } from 'vscode';
import * as stBar from './statusBar';
import * as vbs from './vbsModule';
import * as vbecmCommon from './vbecmCommon';
import { DiffFileInfo, FileDiff, fileDiffProvider } from './diffFiles';
import * as vbeOutput from './vbeOutput';

const STRING_EMPTY = '';

const testAndConfirm =
  (message: string) =>
  async (
    baseDir: string,
    srcDir: string,
    vbeDir: string,
    compareTo: 'src' | 'vbe',
    moduleFileNames: string[]
  ) => {
    const ans = await vbs.comparePathAndGo(
      baseDir,
      srcDir,
      vbeDir,
      compareTo,
      message,
      moduleFileNames
    );
    return ans;
  };

// for book
// export all modules
export async function handlerExportModulesFromFolder(uriFolder: vscode.Uri) {
  const bookPath = await getExcelPathSrcFolder(uriFolder);
  if (bookPath === undefined) {
    vbeOutput.showWarn(uriFolder.fsPath + ' is not a VBE project folder.');
    return;
  }
  await handlerExportModulesFromBook(vscode.Uri.file(bookPath));
}

export async function handlerExportModulesFromBook(uriBook: vscode.Uri) {
  vbeOutput.showInfo(`Start export modules: ${uriBook.fsPath}`);
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
    r && vbeOutput.showInfo('Success export modules.');

    await updateModificationBySystem(uriBook.fsPath);
  } catch (e) {
    vbeOutput.showError('Error export modules.');
    vbeOutput.showError(e);
  } finally {
    displayMenus();
  }
}


export async function handlerCompileFolder(uriFolder: vscode.Uri) {
  const bookPath = await getExcelPathSrcFolder(uriFolder);
  if (bookPath === undefined) {
    vbeOutput.showWarn(uriFolder.fsPath + ' is not a VBE project folder.');
    return;
  }
  await handlerCompile(vscode.Uri.file(bookPath));
}

export function handlerCompile(uriBook: vscode.Uri) {
  vbeOutput.showInfo(`Start compile: ${uriBook.fsPath}`);
  displayMenus(false);
  try {
    vbs.compile(uriBook.fsPath);
    vbeOutput.showInfo('Success compile, may be.');
  } catch (e) {
    vbeOutput.showInfo('Error compile.');
    vbeOutput.showError(e);
  } finally {
    displayMenus(true);
  }
}



// import all modules
export async function handlerImportModulesToBook(uriBook: vscode.Uri) {
  // temp vs vase
  vbeOutput.showInfo(`Start import modules: ${uriBook.fsPath}`);
  const diffTestAndConfirm = testAndConfirm(
    'Excel Vba may be modified. Do you import? Check vbe view.'
  );

  displayMenus(false);
  try {
    const r = await vbs.importModules(uriBook.fsPath, STRING_EMPTY, diffTestAndConfirm);
    r && vbeOutput.showInfo('Success import modules.');

    await updateModificationBySystem();
  } catch (e) {
    vbeOutput.showError('Error import modules.');
    vbeOutput.showError(e);
  } finally {
    displayMenus(true);
  }
}

export async function handlerExportFrxModulesFromBook(uriBook: vscode.Uri) {
  vbeOutput.showInfo(`Start export frx modules: ${uriBook.fsPath}.`);
  displayMenus(false);
  try {
    await vbs.exportFrxModules(uriBook.fsPath);
    vbeOutput.showInfo('Success export frx modules.');
  } catch (e) {
    vbeOutput.showError('Error export frx modules.');
    vbeOutput.showError(e);
  } finally {
    displayMenus();
  }
}

// for module
// export a module
export async function handlerPullModuleAsync(uriModule: vscode.Uri) {
  vbeOutput.showInfo(`Start pull a module: ${uriModule.fsPath}`);
  // if the module file does not exist, return STRING_EMPTY
  const bookPath = await getExcelPathFromModule(uriModule);
  if (bookPath === undefined) {
    vbeOutput.showWarn(uriModule.fsPath + ' is not a VBE project file.');
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
    r && vbeOutput.showInfo('Success pull.');

    await updateModificationBySystem();
  } catch (e) {
    vbeOutput.showError('Error pull.');
    vbeOutput.showError(e);
  } finally {
    displayMenus(true);
  }
}

// commit a module
export async function handlerCommitModuleFromFile(uriModule: vscode.Uri) {
  vbeOutput.showInfo(`Start push a module: ${uriModule.fsPath}`);
  const diffTestAndConfirm = testAndConfirm('Excel Vba may be modified. Do you push this file?');
  const bookPath = await getExcelPathFromModule(uriModule);
  if (bookPath === undefined) {
    vbeOutput.showWarn(uriModule.fsPath + ' is not a VBE project file.');
    return;
  }

  displayMenus(false);
  try {
    const modulePath = uriModule.fsPath;
    const r = await vbs.importModules(bookPath, modulePath, diffTestAndConfirm);
    r && vbeOutput.showInfo('Success push.');

    await updateModificationBySystem();
  } catch (e) {
    vbeOutput.showError('Error push.');
    vbeOutput.showError(e);
  } finally {
    displayMenus(true);
  }
}

export async function handlerGotoVbe(
  textEditor: TextEditor,
  edit: vscode.TextEditorEdit,
  uriModule: vscode.Uri
) {
  displayMenus(false);
  // excel file path
  const bookPath = await getExcelPathFromModule(uriModule);
  if (bookPath === undefined) {
    vbeOutput.showWarn(uriModule.fsPath + ' is not a VBE project file.');
    displayMenus(true);
    return;
  }

  // get module name and line of code
  const activeLineNo = textEditor.selection.active.line;
  const moduleFileName = path.basename(uriModule.fsPath);
  const moduleName = moduleFileName.split('.')[0];
  vbeOutput.showInfo(`Start goto vbe.: ${moduleName}, line: ${activeLineNo}`, false);
  displayMenus(false);
  try {
    const modulePath = uriModule.fsPath;
    await vbs.gotoVbe(bookPath, moduleName, activeLineNo);
    vbeOutput.showInfo('Success goto vbe.', false);
  } catch (e) {
    vbeOutput.showError('Error goto vbe.');
    vbeOutput.showError(e);
  } finally {
    displayMenus(true);
  }
}

export async function handlerVbaRunFromEditor(
  textEditor: TextEditor,
  edit: vscode.TextEditorEdit,
  uriModule: vscode.Uri
) {
  vbeOutput.showInfo('Start run a sub function.');
  displayMenus(false);
  // excel file path
  const bookPath = await getExcelPathFromModule(uriModule);
  if (bookPath === undefined) {
    vbeOutput.showWarn(uriModule.fsPath + ' is not a VBE project file.');
    displayMenus(true);
    return;
  }

  // get sub function
  const activeLine = textEditor.document.lineAt(textEditor.selection.active.line).text;
  const vbaSub = activeLine.match(/Sub (.*)\(\s*\)/i);
  const funcName = vbaSub !== null && vbaSub.length === 2 && vbaSub[1];
  if (funcName === false) {
    vbeOutput.showError('Error run. No sub name.');
    displayMenus(true);
    return;
  }

  vbeOutput.showInfo(`The sub function is ${funcName}`, false);

  displayMenus(false);
  try {
    const modulePath = uriModule.fsPath;
    await vbs.vbaSubRun(bookPath, path.basename(modulePath), funcName);
    vbeOutput.showInfo('Success run, may be.');
  } catch (e) {
    vbeOutput.showError('Error run.');
    vbeOutput.showError(e);
  } finally {
    displayMenus(true);
  }
}

// commit all modules for commit
export async function handlerCommitAllModuleFromFolder(uriFolder: vscode.Uri) {
  vbeOutput.showInfo(`Start push all. ${uriFolder.fsPath}`);
  const bookPath = await getExcelPathSrcFolder(uriFolder);
  if (bookPath === undefined) {
    vbeOutput.showWarn(uriFolder.fsPath + ' is not a VBE project folder.');
    return;
  }

  const diffTestAndConfirm = testAndConfirm(
    'Excel Vba may be modified. Do you push all? Check vbe view.'
  );

  displayMenus(false);
  try {
    const r = await vbs.importModules(bookPath, STRING_EMPTY, diffTestAndConfirm);
    r && vbeOutput.showInfo('Success push all.');
    await updateModificationBySystem();
  } catch (e) {
    vbeOutput.showError('Error push all.');
    vbeOutput.showError(e);
  } finally {
    displayMenus(true);
  }
}

export async function handlerCheckModifiedOnSave(uri: vscode.Uri) {
  vbeOutput.showInfo(`Start check modification on save: ${uri.fsPath}`, false);
  displayMenus(false);
  try {
    const bookPathFromModule = await getExcelPathFromVbeModule(uri);
    const bookPath = bookPathFromModule || (await getExcelPathSrcFolder(uri));
    await updateModificationBySystem(bookPath);
    vbeOutput.showInfo('Success check modification on save. if setting is on.', false);
  } catch (e) {
    vbeOutput.showError('Error CheckModified.');
    vbeOutput.showError(e);
  } finally {
    displayMenus();
  }
}

export async function handlerCheckModifiedOnFolder(uri: vscode.Uri) {
  vbeOutput.showInfo(`Start check modification: ${uri.fsPath}`, false);
  displayMenus(false);
  try {
    const bookPathFromModule = await getExcelPathFromVbeModule(uri);
    const bookPath = bookPathFromModule || (await getExcelPathSrcFolder(uri));
    await updateModificationForce(bookPath);
    vbeOutput.showInfo('Success check modification', false);
  } catch (e) {
    vbeOutput.showError('Error CheckModified.');
    vbeOutput.showError(e);
  } finally {
    displayMenus();
  }
}

// vbe diff tree view
export async function handlerUpdateModificationOnFolder() {
  vbeOutput.showInfo(`Start check modification: refresh button.`, false);
  displayMenus(false);
  try {
    await updateModificationForce();
    vbeOutput.showInfo('Success check modification', false);
  } catch (error) {
    vbeOutput.showError('error update modification');
    vbeOutput.showError(error);
  } finally {
    displayMenus();
  }
}

export async function handlerCommitModuleFromVbeDiff(fileInfo: DiffFileInfo) {
  const bookPath = await getExcelPathFromModule(vscode.Uri.file(fileInfo.compareFilePath!));
  if (bookPath === undefined) {
    vbeOutput.showWarn(fileInfo.moduleName + ' is not a VBE project file.');
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
  displayMenus(false);
  try {
    await vbs.removeModuleSync(bookPath, modulePath);
    await updateModificationBySystem(bookPath);
    vbeOutput.showInfo('Success push(remove module).');
  } catch (error) {
    vbeOutput.showError('error push(remove module)');
    vbeOutput.showError(error);
  } finally {
    displayMenus();
  }
  return true;
}

export async function handlerResolveVbeConflicting(fileInfo: DiffFileInfo) {
  vbeOutput.showInfo('Start Resolve Vbe Conflicting', false);

  const bookPath = await getExcelPathFromVbeModule(vscode.Uri.file(fileInfo.compareFilePath!));
  if (bookPath === undefined) {
    vbeOutput.showWarn(fileInfo.compareFilePath + ' is not a VBE project file.');
    return;
  }

  displayMenus(false);
  try {
    await vbs.resolveVbeConflicting(bookPath, fileInfo.moduleName!);
    await updateModificationBySystem();
    vbeOutput.showInfo('Success Resolve Vbe Conflicting', false);
  } catch (error) {
    vbeOutput.showError('Error CheckModified.');
    vbeOutput.showError(error);
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
    vbeOutput.showWarn(resource.compareFilePath + ' is not a VBE project file.');
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

async function updateModificationBySystem(bookPath?: string) {
  const isAutoRefreshDiff = vscode.workspace
    .getConfiguration('vbecm')
    .get<boolean>('AutoRefreshDiff');
  isAutoRefreshDiff && (await updateModificationForce(bookPath));
}

async function updateModificationForce(bookPath?: string) {
  await fileDiffProvider.updateModification(bookPath);
}
