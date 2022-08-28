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


/**
 * get getVbecmDirs
 * @param uriBook
 * @returns 
 */
export function getVbecmDirs(xlsmPath: string, suffix: string = '') {
  const fileDir = path.dirname(xlsmPath);
  const baseName = path.basename(xlsmPath);
  const srcDir = path.resolve(fileDir, 'src_' + baseName);
  const baseDir = path.resolve(srcDir, '.base');
  const tempDir = path.resolve(fileDir, 'src_' + uuidv4() + '_' + suffix);

  return {
    xlsmPath,
    srcDir,
    baseDir,
    tempDir
  };
}


/**
 * 
 */
export type TestConfirm = { (dir1: string, dir2: string, diffTitle: string): Promise<boolean>; };


/**
 * 
 * @param pathBook 
 * @param testConfirm baseDir vs srcDir
 * @returns 
 */
export async function exportModulesToScrAndBase(pathBook: string, testConfirm?:TestConfirm ) : Promise<boolean|undefined>{
  const {srcDir, baseDir, tempDir} = getVbecmDirs(pathBook,'exportModulesToScrAndBase');
    
  // export to uuid folder
  try {
    await exportModuleAsync(pathBook, tempDir, '');  
  } catch (error) {
    console.log(error);
    await rmDirIfExist(tempDir, { recursive: true, force: true }); 
    throw(error);
  }
 
  // test already exported
  const isExistSrc = await dirExists(srcDir);

  //
  //tar.tarToBase(baseDir);

  // check base and srcDir, some file in srcDir are modified?
  if (isExistSrc && testConfirm){
    const ans = await testConfirm(baseDir, srcDir, 'base,src');
    if (ans === false){
      await rmDirIfExist(tempDir, { recursive: true, force: true });
      return false;
    }
  }

  // copy temp to real folder
  try {
    deleteModulesInSrc(srcDir);
    await fse.copy(tempDir, srcDir);
    deleteModulesInSrc(baseDir);
    await fse.copy(tempDir, baseDir);
    // tar.baseToTar(baseDir, false);
  }
  catch (e) {
    throw(e);
  }
  finally{
    await rmDirIfExist(tempDir, { recursive: true, force: true }); 
  }
  return true;
};

/**
 * 
 * @param xlsmPath 
 * @param modulePath 
 * @param testConfirm   baseFile vs srcFile
 * @returns 
 */
 export async function updateModule(xlsmPath: string, modulePath:string, testConfirm? :TestConfirm) : Promise<boolean| undefined> {
  // old version, checkout

  // if the file does not exist, return ''
  if (!fileExists(xlsmPath)) {
    throw(Error(`Excel file does not exist to update src.: ${modulePath}`));
  }


  //module in excel modified?
  const {srcDir, baseDir, tempDir} = getVbecmDirs(xlsmPath,'updateModule');
  const moduleBase = path.basename(modulePath);
  const doTest = testConfirm !== undefined;
  const goNext = !doTest || await testConfirm(path.resolve(baseDir, moduleBase), path.resolve(srcDir, moduleBase), '');
  if (goNext === false){
    await rmDirIfExist(tempDir, { recursive: true, force: true });
    return;
  }

  // export and copy
  try{
    await exportModuleAsync(xlsmPath, srcDir, path.resolve(srcDir, moduleBase));
    await fse.copy(path.resolve(srcDir, moduleBase), path.resolve(baseDir, moduleBase));
    // tar.baseToTar(baseDir, false);
  }
  catch(e)
  {
    throw(e);
  }
  finally
  {
    await rmDirIfExist(tempDir, { recursive: true, force: true });
  }
  return true;
}


/**
 * 
 * @param pathBook 
 * @param testConfirm tempDir vb baseDire
 * @returns 
 */
export async function importModules(pathBook: string, testConfirm?:TestConfirm) :Promise<boolean|undefined> {

  const { xlsmPath, srcDir, baseDir, tempDir } = getVbecmDirs(pathBook,'importModules');

  const isSrcExist = await common.dirExists(srcDir);
  if (!isSrcExist) {
    throw(new Error (`Source folder does not exist: ${srcDir}`));
  }

  const xlsExists = await common.fileExists(xlsmPath);
  if (!xlsExists) {
    throw(new Error (`Excel file does not exist: ${xlsmPath}`));
  }

  // export to uuid folder, for check modified
  try{
    await exportModuleAsync(pathBook, tempDir, '');
  }
  catch(e)
  {
    await rmDirIfExist(tempDir, { recursive: true, force: true }); 
    throw(e);
  }

  // check tempDir and srcDir
  const existBase = await  dirExists(baseDir);
  const isDefinedTestFunc = !!testConfirm;
  const isConfirm = existBase && isDefinedTestFunc && await testConfirm(baseDir, tempDir , '');
  const isGoForward = !isDefinedTestFunc || !existBase || isConfirm;
  await rmDirIfExist(tempDir, { recursive: true, force: true });
  if (isGoForward === false){
    return;
  }
  
  // import main
  try {
    await importModuleSync(pathBook);
    await exportModulesToScrAndBase(pathBook);
  }
  catch (e) {
    throw(e);
  }
  return true;
}
/**
 * 
 * @param xlsmPath 
 * @param modulePath 
 * @param testConfirm vbe module vs base module
 * @returns 
 */
 export async function commitModule(xlsmPath: string, modulePath:string, testConfirm?:TestConfirm) : Promise<boolean| undefined> {
  if (!fileExists(xlsmPath)) {
    throw(Error(`Excel file does not exist to commit.: ${modulePath}`));
  }

  //module in excel modified?
  const {srcDir, baseDir, tempDir} = getVbecmDirs(xlsmPath, 'commitModule');
  try{
    await exportModuleAsync(xlsmPath, tempDir, '');
  }
  catch(e)
  {
    await rmDirIfExist(tempDir, { recursive: true, force: true }); 
    throw(e);
  }

  // check tempDir and srcDir
  const moduleBase = path.basename(modulePath);
  const ext = path.extname(moduleBase).toLocaleLowerCase();
  const moduleFrx = ext === '.frm' ? moduleBase.slice(ext.length) + '.frx': '';


  const doTest = testConfirm !== undefined;
  const goNext = !doTest || await testConfirm(path.resolve(tempDir, moduleBase), path.resolve(baseDir, moduleBase), '');
  await rmDirIfExist(tempDir, { recursive: true, force: true });
  if (goNext === false){
    return;
  }
  try {
    await importModuleSync(xlsmPath, modulePath);
    await fse.copy(path.resolve(srcDir, moduleBase), path.resolve(baseDir, moduleBase));
    moduleFrx && await fse.copy(path.resolve(srcDir, moduleFrx), path.resolve(baseDir, moduleFrx));    
  } catch (error) {
    throw(error);
  }
  return true;
}

export async function exportFrxModules(pathBook: string){

  if (!fileExists(pathBook)){
    throw(Error(`Excel file does not exist.: ${pathBook}`));
  }

  const {srcDir, baseDir, tempDir} = getVbecmDirs(pathBook,'exportFrxModules');

  // export to uuid folder
  try{
    await exportModuleAsync(pathBook, tempDir, '');
  }
  catch(e)
  {
    await rmDirIfExist(tempDir, { recursive: true, force: true }); 
    throw(e);
  }

  const fileNameList = fse.readdirSync(tempDir);
  const targetFileNames = fileNameList.filter(RegExp.prototype.test, /.*\.frx$/i);

  targetFileNames.forEach(fileName => {
    fse.copyFileSync(path.resolve(tempDir,fileName),path.resolve(srcDir,fileName));
    fse.copyFileSync(path.resolve(tempDir,fileName),path.resolve(baseDir,fileName));
    // tar.baseToTar(baseDir, false);
  });

  await rmDirIfExist(tempDir, { recursive: true, force: true });
}

export async function vbaSubRun(xlsmPath: string, modulePath: string, funcName: string) {
  // excel file path
  if (!fileExists(xlsmPath)){
    throw(Error(`Excel file does not exist to run.: ${modulePath}`));
  }

  // sub function
  if (!funcName){
    throw(Error('No sub function is selected.'));
  }

  const moduleName = path.basename(modulePath, path.extname(modulePath));

  const { err, status, retValue } = runVbs('runVba.vbs', [xlsmPath, moduleName + '.' +funcName]);
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

export async function comparePathAndGo(base: string, target: string, diffTitle: string, confirmMessage: string) {

  const r = comparePath(base, target);

  if (r.differences > 0) {
    const result = r.diffSet!.filter(_ => _.state !== 'equal').map(_ => {
      return (_.name1 || _.name2 || '') + ' : ' + _.reason;
    });
    vbeOutput.clear();
    vbeOutput.appendLine(`======== ${diffTitle} ===========>`);
    vbeOutput.appendLine(result.join('\n') || '');

    const ans = await vscode.window.showInformationMessage(confirmMessage, 'Yes', 'No');

    vbeOutput.clear();
    return ans;
  }
  return 'Yes';
}


export function comparePath(path1: string, path2: string){
  const options: dirCompare.Options = {
    //compareSize: true,  // comment out for not detecting empty line diff.
    compareContent: true,
    compareFileSync: dirCompare.fileCompareHandlers.lineBasedFileCompare.compareSync,
    compareFileAsync: dirCompare.fileCompareHandlers.lineBasedFileCompare.compareAsync,
    skipSubdirs: true,
    includeFilter:'*.cls,*.bas,*.frm',
    excludeFilter: '**/*.frx,.base,.current,.git,.gitignore',
    ignoreEmptyLines: true
  };

  const r = dirCompare.compareSync(path1, path2, options);

  return r;
}

export async function deleteModulesInSrc(pathSrc: string){
  try {

    const isExist = await common.dirExists(pathSrc);
    if (!isExist)
    {
      // no folder
      return;
    }

    const files = fse.readdirSync(pathSrc);
    
    // delete only vba module file
    const moduleList = files.forEach((file) => {
        const fullPath = path.resolve(pathSrc, file);
        const isFile = fse.statSync(fullPath).isFile();
        const ext  = path.extname(file).toLowerCase();
        if (isFile && ['.bas','.frm','.cls', '.frx'].includes(ext)){
          fse.rmSync(fullPath);
        }
    });   
  } catch (error) {
    throw(error);
  }
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




async function exportModuleAsync(xlsmPath: string, pathToExport: string, moduleName: string) {

  const moduleCountInBook = getFilesCountToExport(xlsmPath);

  const isSheet =  'True';
  const isForm = 'True';
  // export
  const { err, status } = runVbs('export.vbs', [xlsmPath, pathToExport, moduleName, isSheet, isForm]);
  if (status !== 0 || err) {
    const error = new Error(err);
    throw (error);
  }

  const moduleCountInfolder = await getModulesCountInFolder(pathToExport);

  // test exported files
  // todo moduleName is set, more check
  if (moduleCountInBook !== moduleCountInfolder) {
    throw (Error('Vbe module count is differ form the count of folders'));
  }
}


async function importModuleSync(pathBook: string, modulePath: string = '') {
  const { err, status } = runVbs('import.vbs', [pathBook, modulePath]);

  if (status !== 0 || err) {
    const error = Error(err);
    throw (error);
  }

  const { srcDir, tempDir } = getVbecmDirs(pathBook, 'importModule');

  try {
    await exportModuleAsync(pathBook, tempDir, '');
    const r = comparePath(srcDir, tempDir);
    if (r.same === false){
      throw(Error('Import verify Error.'));
    }
  } catch (error) {
    console.log(error);
    throw(error);
  }
  finally{
    await rmDirIfExist(tempDir, { recursive: true, force: true }); 
  }
}


type RunVbs = {err: string, out: string, retValue: string, status: number | boolean| null};
/**
 * 
 * @param param vbs, ...params
 * @returns messages
 */
function runVbs(script: string, param: string[]) {
  try {
    const scriptPath = path.resolve(getVbsPath(), script);
    const vbs = spawnSync('cscript.exe', ['//Nologo', scriptPath, ...param]);
    const err = s2u(vbs.stderr);
    const out = s2u(vbs.stdout);
    const retValue = out.split('\r\n').slice(-2)[0];
    common.myLog(`====================vbs run in=======================`);
    common.myLog(scriptPath, `Script`);
    common.myLog(err, `stderr`);
    common.myLog(out, `stdout`);
    common.myLog(retValue, `retVal`);
    common.myLog(`${vbs.status}`, `status`);
    common.myLog(`====================vbs run out=======================`);
    return { err, out, retValue, status: vbs.status };
  }
  catch (e: unknown) {
    const errMessage = (e instanceof Error) ? e.message : 'vbs run error.';
    return { err: errMessage, out: '', retValue: '', status: 10 };
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


export async function rmDirIfExist(pathFolder: string, option: {}){
  try {

    const isExist = await common.dirExists(pathFolder);
    if (!isExist)
    {
      // no folder no delete
      return;
    }
    await fse.promises.rm(pathFolder, option);

  } catch (error) {
    throw(error);
  }
}

