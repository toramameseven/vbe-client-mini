import * as path from 'path';
import * as vscode from 'vscode';
import * as fse from 'fs-extra';
import * as iconv from 'iconv-lite';
import { v4 as uuidv4 } from 'uuid';
import { spawnSync } from 'child_process';
import dirCompare = require('dir-compare');
import { dirExists, fileExists } from './vbecmCommon';
import * as vbecmCommon from './vbecmCommon';
import { DiffFileInfo, fileDiffProvider } from './diffFiles';
import { vbeReadOnlyDocumentProvider } from './extension';
import * as vbeOutput from './vbeOutput';

export const FOLDER_VBS = 'vbs';
export const FOLDER_PREFIX_SRC = 'src_';
export const FOLDER_BASE = '.base';
export const FOLDER_VBE = '.vbe';
export const STRING_EMPTY = '';

export type TestConfirm = {
  (
    baseDir: string,
    srcDir: string,
    vbeDir: string,
    compareTo: 'src' | 'vbe',
    moduleFileNames: string[]
  ): Promise<boolean>;
};

export function getVbecmDirs(bookPath: string, suffix: string = STRING_EMPTY) {
  const fileDir = path.dirname(bookPath);
  const baseName = path.basename(bookPath);
  const srcDir = path.resolve(fileDir, FOLDER_PREFIX_SRC + baseName);
  const baseDir = path.resolve(srcDir, FOLDER_BASE);
  const vbeDir = path.resolve(srcDir, FOLDER_VBE);
  const tempDir = path.resolve(fileDir, 'src_' + uuidv4() + '_' + suffix);

  return {
    bookPath,
    srcDir,
    baseDir,
    vbeDir,
    tempDir,
  };
}

export async function exportModulesAndSynchronize(
  pathBook: string,
  moduleFileNameOrPath: string = STRING_EMPTY,
  testConfirm?: TestConfirm
): Promise<boolean | undefined> {
  const { srcDir, baseDir, vbeDir } = getVbecmDirs(pathBook, 'exportModulesAndSynchronize');

  // get base name
  const moduleFileName = path.basename(moduleFileNameOrPath);

  // test already exported
  const isExistSrc = await dirExists(srcDir);

  // export modules to vbe folder
  try {
    const existSrcFile = await vbecmCommon.fileExists(srcDir);
    if (existSrcFile) {
      throw Error(`file ${srcDir} exists. can not create folder`);
    }
    fse.mkdirSync(srcDir, { recursive: true });
    // all module are update in vbeDir
    await exportModuleAsync(pathBook, vbeDir, STRING_EMPTY);
  } catch (error) {
    console.log(error);
    throw error;
  }

  // check base and srcDir, some file in srcDir are modified?
  if (isExistSrc && testConfirm) {
    const ans = await testConfirm(baseDir, srcDir, vbeDir, 'src', [moduleFileName]);
    if (ans === false) {
      return false;
    }
  }

  // copy vbe to src and base
  try {
    if (moduleFileName === STRING_EMPTY) {
      deleteModulesInFolder(srcDir);
      await fse.copy(vbeDir, srcDir);
      deleteModulesInFolder(baseDir);
      await fse.copy(vbeDir, baseDir);
    }
    if (moduleFileName !== STRING_EMPTY) {
      copyModuleSync(moduleFileName, vbeDir, srcDir);
      copyModuleSync(moduleFileName, vbeDir, baseDir);
    }
  } catch (e) {
    throw e;
  }
  return true;
}

export async function importModules(
  pathBook: string,
  modulePathOrFile: string,
  testConfirm?: TestConfirm
): Promise<boolean | undefined> {
  const { bookPath, srcDir, baseDir, vbeDir } = getVbecmDirs(pathBook, 'importModules');

  const moduleFileName = path.basename(modulePathOrFile);

  const isSrcExist = await vbecmCommon.dirExists(srcDir);
  if (!isSrcExist) {
    throw new Error(`Source folder does not exist: ${srcDir}`);
  }

  const xlsExists = await vbecmCommon.fileExists(bookPath);
  if (!xlsExists) {
    throw new Error(`Excel file does not exist: ${bookPath}`);
  }

  // export to vbe folder, if module in vbe is modified.  vbe vs. base
  try {
    await exportModuleAsync(pathBook, vbeDir, STRING_EMPTY);
  } catch (e) {
    throw e;
  }

  // check vbeDir and srcDir
  const existBase = await dirExists(baseDir);
  const isDefinedTestFunc = !!testConfirm;
  const isConfirm =
    existBase &&
    isDefinedTestFunc &&
    (await testConfirm(baseDir, srcDir, vbeDir, 'vbe', [moduleFileName]));
  const isGoForward = !isDefinedTestFunc || !existBase || isConfirm;
  if (isGoForward === false) {
    return;
  }

  // import main
  let successImport = false;
  try {
    // this is import
    await importModuleSync(pathBook, moduleFileName);
    successImport = true;
  } catch (e) {
    successImport = false;
    vbeOutput.showWarn('import error. now try to recover.', true);
    vbeOutput.showWarn(e, true);
  }

  // import recovery
  try {
    if (!successImport) {
      await importModuleSync(pathBook, STRING_EMPTY, true);
      // for recovery not update modules, so return
      return false;
    }
  } catch (e) {
    throw Error('Fail recover!!');
  }

  try {
    // synchronize vbe to src_dir
    await exportModulesAndSynchronize(pathBook, moduleFileName);
  } catch (e) {
    throw e;
  }
  return true;
}

export async function exportFrxModules(pathBook: string) {
  if (!(await fileExists(pathBook))) {
    throw Error(`Excel file does not exist.: ${pathBook}`);
  }

  const { srcDir, baseDir, vbeDir, tempDir } = getVbecmDirs(pathBook, 'exportFrxModules');

  // export to uuid folder
  try {
    await exportModuleAsync(pathBook, tempDir, STRING_EMPTY);
  } catch (e) {
    await vbecmCommon.rmDirIfExist(tempDir, { recursive: true, force: true });
    throw e;
  }

  const fileNameList = fse.readdirSync(tempDir);
  const targetFileNames = fileNameList.filter(RegExp.prototype.test, /.*\.frx$/i);

  targetFileNames.forEach((fileName) => {
    copyModuleSync(fileName, tempDir, srcDir);
    copyModuleSync(fileName, tempDir, vbeDir);
    copyModuleSync(fileName, tempDir, baseDir);
  });
  await vbecmCommon.rmDirIfExist(tempDir, { recursive: true, force: true });
}

export async function vbaSubRun(bookPath: string, modulePath: string, funcName: string) {
  // excel file path
  if (!(await fileExists(bookPath))) {
    throw Error(`Excel file does not exist to run.: ${modulePath}`);
  }

  // sub function
  if (!funcName) {
    throw Error('No sub function is selected.');
  }

  const moduleFileName = path.basename(modulePath, path.extname(modulePath));

  const { err, status, retValue } = runVbs('runVba.vbs', [
    bookPath,
    moduleFileName + '.' + funcName,
  ]);
  if (status !== 0 || err) {
    throw Error(err);
  }
  return retValue;
}

export async function gotoVbe(bookPath: string, moduleName: string, codeLineNo: number) {
  // excel file path
  if (!(await fileExists(bookPath))) {
    throw Error(`Excel file does not exist to go.: ${bookPath}`);
  }

  const { err, status, retValue } = runVbs('selectModuleLine.vbs', [
    bookPath,
    moduleName,
    codeLineNo.toString(),
  ]);
  if (status !== 0 || err) {
    throw Error(err);
  }
  return retValue;
}

export function compile(bookPath: string) {
  const { err, status } = runVbs('compile.vbs', [bookPath]);
  if (status !== 0 || err) {
    throw Error(err);
  }
}

export function getFilesCountToExport(bookPath: string) {
  const { err, status, retValue } = runVbs('getModules.vbs', [bookPath]);
  if (status !== 0 || err) {
    throw Error(err);
  }
  return isNumeric(retValue) ? Number(retValue) : -1;

  function isNumeric(val: string) {
    return /^-?\d+$/.test(val);
  }
}

export function closeBook(bookPath: string) {
  // get path, but close a same name book.
  const { err, status } = runVbs('closeBook.vbs', [bookPath]);
  if (status !== 0 || err) {
    throw Error(err);
  }
}

export async function comparePathAndGo(
  base: string,
  src: string,
  vbe: string,
  diffTo: 'src' | 'vbe',
  confirmMessage: string,
  moduleFilePaths: string[]
) {
  const moduleFileName = moduleFilePaths ? path.basename(moduleFilePaths[0]) : STRING_EMPTY;
  const pathToCompare = (pPath: string) =>
    moduleFileName ? path.resolve(pPath, moduleFileName) : pPath;

  const basePath = pathToCompare(base);
  const srcPath = pathToCompare(src);
  const vbePath = pathToCompare(vbe);

  const toDiff = diffTo === 'vbe' ? vbePath : srcPath;
  const doDiff = (await fileExists(toDiff)) || (await dirExists(toDiff));

  const r = doDiff && comparePath(basePath, toDiff);

  if (r && r.differences > 0) {
    const ans = await modalDialogShow(confirmMessage);
    return ans;
  }
  return true;
}

export type DiffState = 'modified' | 'add' | 'removed' | 'notModified' | 'conflicting';
export type DiffWith = 'DiffWithSrc' | 'DiffWithVbe' | 'DiffConf';
export type DiffTitle = 'base-vbe: ' | 'base-src: ' | 'vbe-src: ';

function createDiffInfo(
  r: dirCompare.Result,
  diffWith: DiffWith,
  dirBase: string,
  dirCompare: string
): DiffFileInfo[] {
  const result = r
    .diffSet!.filter((_) => _.state !== 'equal')
    .map((_) => {
      const moduleName = _.name1 || _.name2 || STRING_EMPTY;
      let diffState: DiffState = 'add';
      if (_.name2 === undefined || _.size2 === 0) {
        diffState = 'removed';
      } else if (_.name1 === _.name2) {
        diffState = 'modified';
      }

      const file1 = path.resolve(dirBase, moduleName);
      const isBase = _.name1 !== undefined;
      const file2 = path.resolve(dirCompare, moduleName);
      const isCompare = _.name2 !== undefined;
      const titleBaseTo: DiffTitle = diffWith === 'DiffWithVbe' ? 'base-vbe: ' : 'base-src: ';
      const obj: DiffFileInfo = {
        moduleName,
        diffState,
        baseFilePath: file1,
        compareFilePath: file2,
        titleBaseTo,
        isBase,
        isCompare,
        diffWith,
      };
      return obj;
    });
  return result;
}

export function createConflictingInfo(d1: DiffFileInfo[], d2: DiffFileInfo[]): DiffFileInfo[] {
  const common = d1.filter((_) => d2.filter((_d2) => _d2.moduleName === _.moduleName).length);
  if (!common.length) {
    return [];
  }

  const conflictingInfo: DiffFileInfo[] = common.map((_d1) => {
    const _d2 = d2.filter((_) => _.moduleName === _d1.moduleName);
    const obj: DiffFileInfo = {
      moduleName: _d1.moduleName,
      diffState: 'conflicting',
      baseFilePath: _d2[0].compareFilePath,
      compareFilePath: _d1.compareFilePath,
      titleBaseTo: 'vbe-src: ',
      isBase: true,
      isCompare: true,
      diffWith: 'DiffConf',
    };
    return obj;
  });
  const r = conflictingInfo.filter((_) => !!_);
  return r;
}

export function comparePath(path1: string, path2: string) {
  const additionalDiffExclude = vscode.workspace
    .getConfiguration('vbecm')
    .get<string>('diffExclude');
  const diffExclude =
    `**/*.frx,.git,.gitignore, .vscode, ${FOLDER_BASE}, ${FOLDER_VBE}` +
    (additionalDiffExclude ? ',' + additionalDiffExclude : '');

  const options: dirCompare.Options = {
    //compareSize: true,  // comment out for not detecting empty line diff.
    compareContent: true,
    compareFileSync: dirCompare.fileCompareHandlers.lineBasedFileCompare.compareSync,
    //compareFileAsync: dirCompare.fileCompareHandlers.lineBasedFileCompare.compareAsync,
    skipSubdirs: true,
    includeFilter: '*.cls,*.bas,*.frm',
    excludeFilter: diffExclude,
    ignoreAllWhiteSpaces: true,
    ignoreEmptyLines: true,
    ignoreCase: true,
    ignoreContentCase: true,
  };

  const r = dirCompare.compareSync(path1, path2, options);
  return r;
}

export async function deleteModulesInFolder(pathSrc: string) {
  try {
    const isExist = await vbecmCommon.dirExists(pathSrc);
    if (!isExist) {
      // no folder
      return;
    }

    const files = fse.readdirSync(pathSrc);
    const extensions = ['.bas', '.frm', '.cls', '.frx'];
    // delete only vba module file
    files.forEach((file) => {
      const fullPath = path.resolve(pathSrc, file);
      const isFile = fse.statSync(fullPath).isFile();
      const ext = path.extname(file).toLowerCase();
      if (isFile && extensions.includes(ext)) {
        fse.rmSync(fullPath, { force: true });
      }
    });
  } catch (error) {
    throw error;
  }
}

export async function compareModules(pathBook: string) {
  const { srcDir, baseDir, vbeDir } = getVbecmDirs(pathBook, 'compareModules');

  await exportModuleAsync(pathBook, vbeDir, STRING_EMPTY);
  const r = comparePath(baseDir, srcDir);
  const r2 = comparePath(baseDir, vbeDir);
  const diffBaseSrc: DiffFileInfo[] = createDiffInfo(r, 'DiffWithSrc', baseDir, srcDir);
  const diffBaseVbe: DiffFileInfo[] = createDiffInfo(r2, 'DiffWithVbe', baseDir, vbeDir);
  const conflicting = createConflictingInfo(diffBaseSrc, diffBaseVbe);
  return { diffBaseSrc, diffBaseVbe, conflicting };
}

async function getModulesCountInFolder(pathSrc: string) {
  try {
    const isExist = await vbecmCommon.dirExists(pathSrc);
    if (!isExist) {
      // no folder
      return -1;
    }

    const files = fse.readdirSync(pathSrc);
    const targetFiles = files.filter((file) =>
      ['.bas', '.frm', '.cls', '.frx'].includes(path.extname(file).toLowerCase())
    );
    return targetFiles.length;
  } catch (error) {
    throw error;
  }
}

export async function resolveVbeConflicting(
  pathBook: string,
  moduleFileName: string,
  autoDialog?: boolean
) {
  const { srcDir, baseDir, vbeDir } = getVbecmDirs(pathBook, 'resolveVbeConflicting');
  const s = path.resolve(srcDir, moduleFileName);
  const v = path.resolve(vbeDir, moduleFileName);
  const b = path.resolve(baseDir, moduleFileName);

  const isVbe = await vbecmCommon.fileExists(v);
  const isBase = await vbecmCommon.fileExists(b);
  const isSrc = await vbecmCommon.fileExists(s);

  if (isBase && isVbe) {
    // VBE modified, copy vbe to base, src is remain now
    const r = await modalDialogShow('Do you copy the module in the vbe to the base.', autoDialog);
    r && copyModuleSync(moduleFileName, vbeDir, baseDir);
  } else if (isBase) {
    // deleted in vbe, modules in base and src will be delete.
    const r = await modalDialogShow(
      'Do you delete the modules in the base and source.',
      autoDialog
    );
    r && fse.rmSync(b, { force: true });
    r && fse.rmSync(s, { force: true });
  } else if (isVbe) {
    // added in vbe, copy vbe to base and src.
    // copy frx ,too
    const r = await modalDialogShow(
      'Do you copy the module in the vbe to the base and the source.',
      autoDialog
    );
    r && copyModuleSync(moduleFileName, vbeDir, baseDir);
    r && copyModuleSync(moduleFileName, vbeDir, srcDir);
  }
}

function copyModuleSync(moduleFile: string, fromDir: string, toDir: string) {
  const fromPath = path.resolve(fromDir, moduleFile);
  const toPath = path.resolve(toDir, moduleFile);
  fse.rmSync(toPath, { force: true });
  fse.copyFileSync(fromPath, toPath);

  // for frx
  // get (.xxx)
  const extension = path.extname(moduleFile).toLocaleLowerCase();
  const frxModuleNameWithOutEx = path.basename(moduleFile).slice(0, -extension.length);
  if (extension !== '.frm') {
    return;
  }
  const fromPathFrx = path.resolve(fromDir, frxModuleNameWithOutEx + '.frx');
  const toPathFrx = path.resolve(toDir, frxModuleNameWithOutEx + '.frx');
  fse.rmSync(toPathFrx, { force: true });
  fse.copyFileSync(fromPathFrx, toPathFrx);
}

/**
 * diff base to src or vbe
 * TODO is src is empty, create new file in src
 * @param resource DiffFileInfo
 */
export async function diffBaseTo(resource: DiffFileInfo): Promise<void> {
  const b = vscode.Uri.file(resource.baseFilePath ?? STRING_EMPTY);
  const readOnlyDocumentBase = await createVirtualModule(b);

  const existFile = await fileExists(resource.compareFilePath ?? STRING_EMPTY);
  const c = vscode.Uri.file(resource.compareFilePath ?? STRING_EMPTY);
  const virtualC = await createVirtualModule(c);

  const documentCompare =
    resource.diffWith === 'DiffWithVbe' ? virtualC : resource.isCompare && existFile ? c : virtualC;

  const t = (resource.titleBaseTo ?? STRING_EMPTY) + resource.moduleName;
  await vscode.commands.executeCommand(
    'vscode.diff',
    readOnlyDocumentBase,
    documentCompare ?? virtualC,
    t
  );
}
/**
 * diff vbe and src
 * @param resource DiffFileInfo
 * @param bookPath
 */
export async function diffSrcToVbe(resource: DiffFileInfo, bookPath: string): Promise<void> {
  const { srcDir, vbeDir } = getVbecmDirs(bookPath);

  const m = resource.moduleName ?? STRING_EMPTY;
  const v = vscode.Uri.file(path.resolve(vbeDir, m));
  const s = vscode.Uri.file(path.resolve(srcDir, m));
  const t = 'vbe-src: ' + resource.moduleName;
  const isExistInSrc = await vbecmCommon.fileExists(path.resolve(srcDir, m));

  const readOnlyDocumentVbe = await createVirtualModule(v);
  const documentSrc = await createVirtualModule(s);
  const srcForDiff = isExistInSrc ? s : documentSrc;

  await vscode.commands.executeCommand('vscode.diff', readOnlyDocumentVbe, srcForDiff, t);
}

async function createVirtualModule(uriModule: vscode.Uri) {
  const uri = vscode.Uri.parse('vbecm:' + uriModule.fsPath);
  const doc = await vscode.workspace.openTextDocument(uri); // calls back into the provider
  vbeReadOnlyDocumentProvider.refresh(uri);
  return doc.uri;
}

async function exportModuleAsync(
  bookPath: string,
  pathToExport: string,
  moduleFileNameOrPath: string
) {
  // delete file then export
  const moduleFileName = path.basename(moduleFileNameOrPath);
  moduleFileName === STRING_EMPTY && (await deleteModulesInFolder(pathToExport));
  const { err, status } = runVbs('export.vbs', [bookPath, pathToExport, moduleFileName]);
  if (status !== 0 || err) {
    const error = new Error(err);
    throw error;
  }

  // verify all modules
  if (moduleFileName === STRING_EMPTY) {
    // modules in book
    const moduleCountInBook = getFilesCountToExport(bookPath);
    // modules in exported folder
    const moduleCountInfolder = await getModulesCountInFolder(pathToExport);
    // test exported files
    // todo moduleFileName is set, more check
    if (moduleCountInBook !== moduleCountInfolder) {
      throw Error(
        `Vbe module count is differ: inBook: ${moduleCountInBook} inSrc: ${moduleCountInfolder}`
      );
    }
    return;
  }

  // verify when moduleFileName defined
  if (!(await vbecmCommon.fileExists(path.resolve(pathToExport, moduleFileName)))) {
    throw Error(`${moduleFileName} is not exported.`);
  }
}

async function importModuleSync(
  pathBook: string,
  modulePath: string = STRING_EMPTY,
  isRecovery: boolean = false
) {
  // get folders
  const { srcDir, baseDir, vbeDir } = getVbecmDirs(pathBook, 'importModule');

  const importDirFrom = isRecovery ? vbeDir : srcDir;

  // do import
  const { err, status } = runVbs('import.vbs', [pathBook, modulePath, importDirFrom]);
  if (status !== 0 || err) {
    throw Error(err);
  }

  if (isRecovery) {
    // do not verify
    return;
  }

  //  verify src and exported(once imported)
  //  comparePath does not work with ignore case.
  // compare dir or file ?
  const moduleFileName = path.basename(modulePath);
  const srcPath = modulePath ? path.resolve(srcDir, moduleFileName) : srcDir;
  const vbePath = modulePath ? path.resolve(vbeDir, moduleFileName) : vbeDir;

  try {
    await exportModuleAsync(pathBook, vbeDir, moduleFileName);
    const r = comparePath(srcPath, vbePath);
    if (r.same === false) {
      throw Error('Import verify Error!! try files.insertFinalNewline is On');
    }
  } catch (error) {
    throw error;
  }
}

export async function removeModuleSync(pathBook: string, modulePath: string) {
  // get folders
  const { srcDir, baseDir, vbeDir } = getVbecmDirs(pathBook, 'importModule');
  // do remove
  const { err, status } = runVbs('remove.vbs', [pathBook, modulePath]);
  if (status !== 0 || err) {
    const error = Error(err);
    throw error;
  }

  // remove module in vbe and base
  const moduleFileName = path.basename(modulePath);
  const basePath = path.resolve(baseDir, moduleFileName);
  const vbePath = path.resolve(vbeDir, moduleFileName);
  await fse.promises.rm(basePath, { force: true });
  await fse.promises.rm(vbePath, { force: true });
}

type RunVbs = { err: string; out: string; retValue: string; status: number | null };
function runVbs(script: string, param: string[]): RunVbs {
  try {
    const rootFolder = path.dirname(__dirname);
    const vbsPath = path.resolve(rootFolder, 'vbs');
    const scriptPath = path.resolve(vbsPath, script);

    const vbs = spawnSync('cscript.exe', ['//Nologo', scriptPath, ...param], { timeout: 60000 });
    const err = s2u(vbs.stderr);
    const out = s2u(vbs.stdout);
    const retValue = out.split('\r\n').slice(-2)[0];
    const vbsName = path.basename(scriptPath);
    vbecmCommon.myLog(`+++++++++++++++++++++ start ${vbsName} +++++++++++++++++++++`);
    vbecmCommon.myLog(scriptPath, `Script`);
    vbecmCommon.myLog(err, `stderr`);
    vbecmCommon.myLog(out, `stdout`);
    vbecmCommon.myLog(retValue, `retVal`);
    vbecmCommon.myLog(`${vbs.status}`, `status`);
    vbecmCommon.myLog(`--------------------- End ${vbsName} ---------------------`);
    return { err, out, retValue, status: vbs.status };
  } catch (e: unknown) {
    const errMessage = e instanceof Error ? e.message : 'vbs run error.';
    return { err: errMessage, out: STRING_EMPTY, retValue: STRING_EMPTY, status: 10 };
  }
}

// shift jis 2 utf8
// only japanese
export function s2u(sb: Buffer) {
  const vbsEncode =
    vscode.workspace.getConfiguration('vbecm').get<string>('vbsEncode') || 'windows-31j';
  return iconv.decode(sb, vbsEncode);
}

async function modalDialogShow(message: string, retValue?: boolean) {
  if (retValue !== undefined) {
    return retValue;
  }
  const ans = await vscode.window.showInformationMessage(
    message,
    { modal: true },
    { title: 'No', isCloseAffordance: true, dialogValue: false },
    { title: 'Yes', isCloseAffordance: false, dialogValue: true }
  );
  return ans?.dialogValue ?? false;
}
