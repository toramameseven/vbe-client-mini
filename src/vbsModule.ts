import * as path from 'path';
import * as vscode from 'vscode';
import * as fse from 'fs-extra';
import * as iconv from 'iconv-lite';
import { vbeOutput } from './vbeOutput';
import { v4 as uuidv4 } from 'uuid';
import {spawnSync} from 'child_process';
import dirCompare = require('dir-compare');
import { dirExists, fileExists } from './common';
import * as common from './common';

export const FOLDER_VBS = 'vbs';
export const FOLDER_PREFIX_SRC = 'src_';
export const FOLDER_BASE = '.base';
export const FOLDER_VBE = '.vbe';
export const EXT_BASE = '.base';
export const EXT_VBE = '.vbe';
export const STRING_EMPTY = '';

export function getVbecmDirs(bookPath: string, suffix: string = STRING_EMPTY) {
  const fileDir = path.dirname(bookPath);
  const baseName = path.basename(bookPath);
  const srcDir = path.resolve(fileDir, FOLDER_PREFIX_SRC + baseName);
  const baseDir = path.resolve(srcDir, FOLDER_BASE);
  const vbeDir = path.resolve(srcDir, FOLDER_VBE);
  const tempDir = path.resolve(fileDir, 'src_' + uuidv4() + '_' + suffix);

  return {
    bookPath, srcDir, baseDir, vbeDir, tempDir
  };
}


export type TestConfirm = { (base: string, src: string, vbe: string, diffTitle: 'src'|'vbe', moduleFileNames?: string[]): Promise<boolean>; };


export async function exportModulesForUpdate(pathBook: string, moduleFineNameOrPath: string = STRING_EMPTY, testConfirm?:TestConfirm ) : Promise<boolean|undefined>{
  const {srcDir, baseDir, vbeDir} = getVbecmDirs(pathBook, 'exportModulesForUpdate');

  // get base name
  const moduleFileName = path.basename(moduleFineNameOrPath);
   
  // test already exported
  const isExistSrc = await dirExists(srcDir);  

  // export modules to vbe folder
  try {
    const existSrcFile = await common.fileExists(srcDir);
    if (existSrcFile) {
      throw(Error(`file ${srcDir} exists. can not create folder`));
    } 
    fse.mkdirSync(srcDir, { recursive: true });
    // all module are update in vbeDir
    await exportModuleAsync(pathBook, vbeDir, STRING_EMPTY);  
  } catch (error) {
    console.log(error);
    throw(error);
  }
 
  // check base and srcDir, some file in srcDir are modified?
  if (isExistSrc && testConfirm){
    const ans = await testConfirm(baseDir, srcDir, vbeDir, 'src', [moduleFileName]);
    if (ans === false){
      return false;
    }
  }

  // copy file or path
  const copySourcePath = moduleFileName ? path.resolve(srcDir, moduleFileName): srcDir;
  const copySourceVbe = moduleFileName ? path.resolve(vbeDir, moduleFileName) : vbeDir;
  const copySourceBase = moduleFileName ? path.resolve(baseDir, moduleFileName) :baseDir;

  // copy src and base
  try {
    moduleFileName === STRING_EMPTY && deleteModulesInFolder(srcDir, true);
    await fse.copy(copySourceVbe, copySourcePath);
    moduleFileName === STRING_EMPTY && deleteModulesInFolder(baseDir, true);
    await fse.copy(copySourceVbe, copySourceBase);

    await common.rmFileIfExist(path.resolve(srcDir, moduleFileName + EXT_BASE), {});
    await common.rmFileIfExist(path.resolve(srcDir, moduleFileName + EXT_VBE), {});
  }
  catch (e) {
    throw(e);
  }
  return true;
};


/* export async function pullModule(bookPath: string, modulePath:string, testConfirm? :TestConfirm) : Promise<boolean| undefined> {
  // if the file does not exist, return STRING_EMPTY
  if (!fileExists(bookPath)) {
    throw(Error(`Excel file does not exist to update src.: ${modulePath}`));
  }

  //module in source is modified?
  const {srcDir, baseDir, vbeDir} = getVbecmDirs(bookPath,'pullModule');
  const moduleBase = path.basename(modulePath);
  const doTest = testConfirm !== undefined;
  const goNext = !doTest || await testConfirm(path.resolve(baseDir, moduleBase), path.resolve(srcDir, moduleBase), path.resolve(vbeDir, moduleBase), 'src');
  if (goNext === false){
    return;
  }

  // export and copy
  try{
    await exportModuleAsync(bookPath, vbeDir, moduleBase); 
    await fse.copy(path.resolve(vbeDir, moduleBase), path.resolve(srcDir, moduleBase));
    await fse.copy(path.resolve(vbeDir, moduleBase), path.resolve(baseDir, moduleBase));
    await common.rmFileIfExist(path.resolve(srcDir, moduleBase + EXT_BASE), {});
    await common.rmFileIfExist(path.resolve(srcDir, moduleBase + EXT_VBE), {});
  }
  catch(e)
  {
    throw(e);
  }
  return true;
} */


export async function importModules(pathBook: string, modulePathOrFile: string, testConfirm?:TestConfirm) :Promise<boolean|undefined> {

  const { bookPath, srcDir, baseDir, vbeDir } = getVbecmDirs(pathBook,'importModules');

  const moduleFileName = path.basename(modulePathOrFile);

  const isSrcExist = await common.dirExists(srcDir);
  if (!isSrcExist) {
    throw(new Error (`Source folder does not exist: ${srcDir}`));
  }

  const xlsExists = await common.fileExists(bookPath);
  if (!xlsExists) {
    throw(new Error (`Excel file does not exist: ${bookPath}`));
  }

  // export to vbe folder, if module in vbe is modified.  vbe vs. base
  try{
    await exportModuleAsync(pathBook, vbeDir, STRING_EMPTY);
  }
  catch(e)
  {
    throw(e);
  }

  // check vbeDir and srcDir
  const existBase = await  dirExists(baseDir);
  const isDefinedTestFunc = !!testConfirm;
  const isConfirm = existBase && isDefinedTestFunc && await testConfirm(baseDir, srcDir, vbeDir, 'vbe');
  const isGoForward = !isDefinedTestFunc || !existBase || isConfirm;
  if (isGoForward === false){
    return;
  }
  
  // import main
  let successImport = false;
  try {
    // this is import
    await importModuleSync(pathBook, moduleFileName);
    successImport = true;
  }
  catch (e) {
    successImport = false;
    vscode.window.showWarningMessage('import error. now try to recover.');
  }

  // import recovery
  try {
    if (!successImport){
      await importModuleSync(pathBook, STRING_EMPTY, true);
      // for recovery not update modules, so return 
      vscode.window.showWarningMessage('Success recover!!');
      return false;
    }
  }
  catch (e) {
    vscode.window.showErrorMessage('Fail recover!!');
    throw(e);
  }

  try {
    // synchronize vbe to src_dir
      await exportModulesForUpdate(pathBook, moduleFileName);
  }
  catch (e) {
    throw(e);
  }
  return true;
}

/* export async function commitModule(bookPath: string, modulePath:string, testConfirm?:TestConfirm) : Promise<boolean| undefined> {
  // dose book exist
  if (!fileExists(bookPath)) {
    throw(Error(`Excel file does not exist to commit.: ${modulePath}`));
  }

  // module in excel modified?
  // so export module(s)
  const {srcDir, baseDir, vbeDir} = getVbecmDirs(bookPath, 'commitModule');
  try{
    await exportModuleAsync(bookPath, vbeDir, modulePath);
  }
  catch(e)
  {
    throw(e);
  }

  // really import(commit)?
  const moduleFileName = path.basename(modulePath);
  const doTest = testConfirm !== undefined;
  const goNext = !doTest || await testConfirm(path.resolve(baseDir, moduleFileName), STRING_EMPTY, path.resolve(vbeDir, moduleFileName), 'vbe');
  if (goNext === false){
    return;
  }

  // import main
  try {
    // now not verify import
    await importModuleSync(bookPath, modulePath);
    await exportModulesForUpdate(bookPath, moduleFileName);  
  } catch (error) {
    throw(error);
  }
  return true;
} */

export async function exportFrxModules(pathBook: string){
  if (!fileExists(pathBook)){
    throw(Error(`Excel file does not exist.: ${pathBook}`));
  }

  const {srcDir, baseDir, vbeDir, tempDir} = getVbecmDirs(pathBook,'exportFrxModules');

  // export to uuid folder
  try{
    await exportModuleAsync(pathBook, tempDir, STRING_EMPTY);
  }
  catch(e) {
    await common.rmDirIfExist(tempDir, { recursive: true, force: true });
    throw(e);
  }

  const fileNameList = fse.readdirSync(tempDir);
  const targetFileNames = fileNameList.filter(RegExp.prototype.test, /.*\.frx$/i);

  targetFileNames.forEach(fileName => {
    fse.copyFileSync(path.resolve(tempDir, fileName),path.resolve(srcDir, fileName));
    fse.copyFileSync(path.resolve(tempDir, fileName),path.resolve(vbeDir, fileName));
    fse.copyFileSync(path.resolve(tempDir, fileName),path.resolve(baseDir, fileName));
  });
  await common.rmDirIfExist(tempDir, { recursive: true, force: true });
}

export async function vbaSubRun(bookPath: string, modulePath: string, funcName: string) {
  // excel file path
  if (!fileExists(bookPath)){
    throw(Error(`Excel file does not exist to run.: ${modulePath}`));
  }

  // sub function
  if (!funcName){
    throw(Error('No sub function is selected.'));
  }

  const moduleFileName = path.basename(modulePath, path.extname(modulePath));

  const { err, status, retValue } = runVbs('runVba.vbs', [bookPath, moduleFileName + '.' +funcName]);
  if (status !== 0 || err) {
    throw(Error(err));
  }
  return retValue;
}


export function compile(bookPath: string) {
  const { err, status } = runVbs('compile.vbs', [bookPath]);
  if (status !== 0 || err) {
    throw(Error(err));
  }
}

export function getFilesCountToExport(bookPath: string) {
  const { err, status, retValue } = runVbs('getModules.vbs', [bookPath]);
  if (status !== 0 || err) {
    throw(Error(err));
  }
  return isNumeric(retValue) ? Number(retValue) : -1;

  function isNumeric(val: string) {
    return /^-?\d+$/.test(val);
  }
}


export function closeBook(bookPath: string){
  // get path, but close a same name book.
  const { err, status } = runVbs('closeBook.vbs', [bookPath]);
  if (status !== 0 || err) {
    throw(Error(err));
  }
}

export async function comparePathAndGo(base: string, src: string, vbe: string, diffTo: 'src'|'vbe', confirmMessage: string, moduleFilePaths?: string[]) {
  
  const moduleFileName = moduleFilePaths ? path.basename(moduleFilePaths[0]) : STRING_EMPTY;
  const pathToCompare = (pPath: string) => moduleFileName ? path.resolve(pPath, moduleFileName) : pPath;

  const basePath =pathToCompare(base);
  const srcPath =pathToCompare(src);
  const vbePath =pathToCompare(vbe);

  const r = comparePath(basePath, diffTo ==='vbe' ? vbePath : srcPath);

  if (r.differences > 0) {
    logToOutputTab(r, diffTo);
    await saveDiffFiles(r, basePath, srcPath, vbePath);
    const ans = await vscode.window.showInformationMessage(confirmMessage, 'Yes', 'No');
    vbeOutput.clear();
    return ans;
  }
  return 'Yes';
}


function logToOutputTab(r: dirCompare.Result, diffTitle: string, doClear : boolean = true){
  const result = r.diffSet!.filter(_ => _.state !== 'equal').map(_ => {
    return (_.name1 || _.name2 || STRING_EMPTY) + ' : ' + _.reason;
  });
  doClear && vbeOutput.clear();
  vbeOutput.appendLine(`======== base - ${diffTitle} ===========>`);
  vbeOutput.appendLine(result.join('\n') || STRING_EMPTY);
}

async function saveDiffFiles(r: dirCompare.Result, base: string, src: string, vbe: string){
  const isDir = await common.dirExists(src);

  const dirSrc = isDir ? src : path.dirname(src);
  const dirBase = isDir ? base : path.dirname(base);
  const dirVbe = isDir ? vbe : path.dirname(vbe);

  for (const diff of r.diffSet!){
    if(diff.state !== 'equal'){
      const fileName = (diff.name1 && diff.name2) ? diff.name1 : STRING_EMPTY;
      const baseFile = path.resolve(dirBase, fileName);
      const vbeFile = path.resolve(dirVbe, fileName);
      await common.fileExists(baseFile) && fse.copyFileSync(baseFile, path.resolve(dirSrc, fileName + EXT_BASE));
      await common.fileExists(vbeFile) && fse.copyFileSync(vbeFile, path.resolve(dirSrc, fileName + EXT_VBE));
    }
  }
}


export function comparePath(path1: string, path2: string){
  const options: dirCompare.Options = {
    //compareSize: true,  // comment out for not detecting empty line diff.
    compareContent: true,
    compareFileSync: dirCompare.fileCompareHandlers.lineBasedFileCompare.compareSync,
    //compareFileAsync: dirCompare.fileCompareHandlers.lineBasedFileCompare.compareAsync,
    skipSubdirs: true,
    includeFilter:'*.cls,*.bas,*.frm',
    excludeFilter: `**/*.frx,.git,.gitignore, ${FOLDER_BASE}, ${FOLDER_VBE}`,
    ignoreAllWhiteSpaces: true,
    ignoreEmptyLines: true,
    ignoreCase: true
  };

  const r = dirCompare.compareSync(path1, path2, options);

  return r;
}

export async function deleteModulesInFolder(pathSrc: string, isDiffFile: boolean = false){
  try {

    const isExist = await common.dirExists(pathSrc);
    if (!isExist)
    {
      // no folder
      return;
    }

    const files = fse.readdirSync(pathSrc);

    const extensions = isDiffFile ? ['.bas', '.frm', '.cls', '.frx', EXT_BASE, EXT_VBE] : ['.bas', '.frm', '.cls', '.frx'];
    
    // delete only vba module file
    const moduleList = files.forEach((file) => {
        const fullPath = path.resolve(pathSrc, file);
        const isFile = fse.statSync(fullPath).isFile();
        const ext  = path.extname(file).toLowerCase();
        if (isFile && extensions.includes(ext)){
          fse.rmSync(fullPath);
        }
    });   
  } catch (error) {
    throw(error);
  }
}


export async function compareModules(pathBook: string){
  const {srcDir, baseDir, vbeDir} = getVbecmDirs(pathBook, 'compareModules');

  await exportModuleAsync(pathBook, vbeDir, STRING_EMPTY); 

  const r = comparePath(baseDir, srcDir);
  logToOutputTab(r, 'src');

  const r2 = comparePath(baseDir, vbeDir);
  logToOutputTab(r2, 'vbe', false);

}

export async function getModulesCountInFolder(pathSrc: string){
  try {

    const isExist = await common.dirExists(pathSrc);
    if (!isExist)
    {
      // no folder
      return -1;
    }

    const files = fse.readdirSync(pathSrc);
    const targetFiles = files.filter(
      file => ['.bas','.frm','.cls', '.frx'].includes(path.extname(file).toLowerCase())
    );
    return targetFiles.length;
  } catch (error) {
    throw(error);
  }
}

async function exportModuleAsync(bookPath: string, pathToExport: string, moduleFileNameOrPath: string) {
  // export modules
  const isSheet =  'True';
  const isForm = 'True';

  // delete file then export
  const moduleFileName = path.basename(moduleFileNameOrPath);
  moduleFileName === STRING_EMPTY && await deleteModulesInFolder(pathToExport, true);
  const { err, status } = runVbs('export.vbs', [bookPath, pathToExport, moduleFileName, isSheet, isForm]);
  if (status !== 0 || err) {
    const error = new Error(err);
    throw (error);
  }

  // verify all modules
  if (moduleFileName === STRING_EMPTY){
    // modules in book
    const moduleCountInBook = getFilesCountToExport(bookPath);
    // modules in exported folder
    const moduleCountInfolder = await getModulesCountInFolder(pathToExport);
    // test exported files
    // todo moduleFileName is set, more check
    if (moduleCountInBook !== moduleCountInfolder) {
      throw (Error(`Vbe module count is differ: inBook: ${moduleCountInBook} inSrc: ${moduleCountInfolder}`));
    }
    return;
  }

  // verify when moduleFileName defined
  if (! await common.fileExists(path.resolve(pathToExport, moduleFileName))){
    throw (Error(`${moduleFileName} is not exported.`)); 
  }
}


async function importModuleSync(pathBook: string, modulePath: string = STRING_EMPTY, isRecovery: boolean = false) {

  // get folders
  const { srcDir, baseDir, vbeDir } = getVbecmDirs(pathBook, 'importModule');

  const importDirFrom = isRecovery ? vbeDir : srcDir;

  // do import
  const { err, status } = runVbs('import.vbs', [pathBook, modulePath, importDirFrom]);
  if (status !== 0 || err) {
    const error = Error(err);
    throw (error);
  }

  if (isRecovery){
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
    if (r.same === false){
      logToOutputTab(r, 'src, vbe');
      await saveDiffFiles(r, baseDir, srcDir, vbeDir);
      vscode.window.showWarningMessage('May be Ok, some files are format in VBE engine. Case or Space.');
      //throw(Error('Import verify Error.'));
    }
  } catch (error) {
    console.log(error);
    throw(error);
  }
}


type RunVbs = {err: string, out: string, retValue: string, status: number | boolean| null};

function runVbs(script: string, param: string[]) {
  try {
    const scriptPath = path.resolve(getVbsPath(), script);
    const vbs = spawnSync('cscript.exe', ['//Nologo', scriptPath, ...param]);
    const err = s2u(vbs.stderr);
    const out = s2u(vbs.stdout);
    const retValue = out.split('\r\n').slice(-2)[0];
    const vbsName = path.basename(scriptPath);
    common.myLog(`+++++++++++++++++++++ start ${vbsName} +++++++++++++++++++++`);
    common.myLog(scriptPath, `Script`);
    common.myLog(err, `stderr`);
    common.myLog(out, `stdout`);
    common.myLog(retValue, `retVal`);
    common.myLog(`${vbs.status}`, `status`);
    common.myLog(`--------------------- End ${vbsName} ---------------------`);
    return { err, out, retValue, status: vbs.status };
  }
  catch (e: unknown) {
    const errMessage = (e instanceof Error) ? e.message : 'vbs run error.';
    return { err: errMessage, out: STRING_EMPTY, retValue: STRING_EMPTY, status: 10 };
  }
}

// shift jis 2 utf8
// only japanese
function s2u(sb: Buffer){
  const vbsEncode = vscode.workspace.getConfiguration('vbecm').get<string>('vbsEncode') || 'windows-31j';
  return iconv.decode(sb, vbsEncode);
};


function getVbsPath(){
  const rootFolder = path.dirname(__dirname);
  const vbsPath = path.resolve(rootFolder, 'vbs');
  return vbsPath;
}

async function deleteModuleFolderIfExist(moduleFolder: string){
  await common.rmDirIfExist(moduleFolder, { recursive: true, force: true });
}
