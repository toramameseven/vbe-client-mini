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
export function getVbecmDirs(bookPath: string, suffix: string = '') {
  const fileDir = path.dirname(bookPath);
  const baseName = path.basename(bookPath);
  const srcDir = path.resolve(fileDir, 'src_' + baseName);
  const baseDir = path.resolve(srcDir, '.base');
  const tempDir = path.resolve(fileDir, 'src_' + uuidv4() + '_' + suffix);

  return {
    bookPath,
    srcDir,
    baseDir,
    tempDir
  };
}


/**
 * 
 */
export type TestConfirm = { (base: string, src: string, vbe: string, diffTitle: 'src'|'vbe', moduleFileNames?: string[]): Promise<boolean>; };


/**
 * export modules from book.
 * @param pathBook 
 * @param testConfirm
 * @returns 
 */
export async function exportModulesToScrAndBase(pathBook: string, moduleFileName: string = '', testConfirm?:TestConfirm ) : Promise<boolean|undefined>{
  const {srcDir, baseDir, tempDir} = getVbecmDirs(pathBook, 'exportModulesToScrAndBase');
    
  // export to uuid folder, all modules or selected module
  try {
    await exportModuleAsync(pathBook, tempDir, moduleFileName);  
  } catch (error) {
    console.log(error);
    await rmDirIfExist(tempDir, { recursive: true, force: true }); 
    throw(error);
  }
 
  // test already exported
  const isExistSrc = await dirExists(srcDir);

  // check base and srcDir, some file in srcDir are modified?
  if (isExistSrc && testConfirm){
    const ans = await testConfirm(baseDir, srcDir, tempDir, 'src', [moduleFileName]);
    if (ans === false){
      await rmDirIfExist(tempDir, { recursive: true, force: true });
      return false;
    }
  }

  // copy file or path
  const copySourcePath = moduleFileName ? path.resolve(srcDir, moduleFileName): srcDir;
  const copySourceVbe = moduleFileName ? path.resolve(tempDir, moduleFileName) : tempDir;
  const copySourceBase = moduleFileName ? path.resolve(baseDir, moduleFileName) :baseDir;

  // copy src and base
  try {
    moduleFileName === '' && deleteModulesInSrc(srcDir, true);
    await fse.copy(copySourceVbe, copySourcePath);
    moduleFileName === '' && deleteModulesInSrc(baseDir, true);
    await fse.copy(copySourceVbe, copySourceBase);
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
 * @param bookPath 
 * @param modulePath 
 * @param testConfirm   baseFile vs srcFile
 * @returns 
 */
 export async function fetchModule(bookPath: string, modulePath:string, testConfirm? :TestConfirm) : Promise<boolean| undefined> {
  // if the file does not exist, return ''
  if (!fileExists(bookPath)) {
    throw(Error(`Excel file does not exist to update src.: ${modulePath}`));
  }


  //module in source is modified?
  const {srcDir, baseDir, tempDir} = getVbecmDirs(bookPath,'fetchModule');
  const moduleBase = path.basename(modulePath);
  const doTest = testConfirm !== undefined;
  const goNext = !doTest || await testConfirm(path.resolve(baseDir, moduleBase), path.resolve(srcDir, moduleBase), path.resolve(tempDir, moduleBase), 'src');
  if (goNext === false){
    await rmDirIfExist(tempDir, { recursive: true, force: true });
    return;
  }

  // export and copy
  try{
    await exportModuleAsync(bookPath, tempDir, ''); 
    await fse.copy(path.resolve(tempDir, moduleBase), path.resolve(srcDir, moduleBase));
    await fse.copy(path.resolve(tempDir, moduleBase), path.resolve(baseDir, moduleBase));
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

  const { bookPath, srcDir, baseDir, tempDir } = getVbecmDirs(pathBook,'importModules');

  const isSrcExist = await common.dirExists(srcDir);
  if (!isSrcExist) {
    throw(new Error (`Source folder does not exist: ${srcDir}`));
  }

  const xlsExists = await common.fileExists(bookPath);
  if (!xlsExists) {
    throw(new Error (`Excel file does not exist: ${bookPath}`));
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
  const isConfirm = existBase && isDefinedTestFunc && await testConfirm(baseDir, srcDir, tempDir, 'vbe');
  const isGoForward = !isDefinedTestFunc || !existBase || isConfirm;
  await rmDirIfExist(tempDir, { recursive: true, force: true });
  if (isGoForward === false){
    return;
  }
  
  // import main
  try {
    // this is import
    // now not verify import
    await importModuleSync(pathBook);
    // do import and export, vbe does format file(case and whitespace)
    // so synchronize after import
    await exportModulesToScrAndBase(pathBook);
  }
  catch (e) {
    throw(e);
  }
  return true;
}
/**
 * 
 * @param bookPath 
 * @param modulePath 
 * @param testConfirm vbe module vs base module
 * @returns 
 */
 export async function commitModule(bookPath: string, modulePath:string, testConfirm?:TestConfirm) : Promise<boolean| undefined> {
  // dose book exist
  if (!fileExists(bookPath)) {
    throw(Error(`Excel file does not exist to commit.: ${modulePath}`));
  }

  // module in excel modified?
  // so export module(s)
  const {srcDir, baseDir, tempDir} = getVbecmDirs(bookPath, 'commitModule');
  try{
    await exportModuleAsync(bookPath, tempDir, modulePath);
  }
  catch(e)
  {
    await rmDirIfExist(tempDir, { recursive: true, force: true }); 
    throw(e);
  }

  // really import(commit)?
  const moduleFileName = path.basename(modulePath);
  const doTest = testConfirm !== undefined;
  const goNext = !doTest || await testConfirm(path.resolve(baseDir, moduleFileName), '', path.resolve(tempDir, moduleFileName), 'vbe');
  await rmDirIfExist(tempDir, { recursive: true, force: true });
  if (goNext === false){
    return;
  }

  // import main
  try {
    // now not verify import
    await importModuleSync(bookPath, modulePath);
    await exportModulesToScrAndBase(bookPath, moduleFileName);  
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
  
  const moduleFileName = moduleFilePaths ? path.basename(moduleFilePaths[0]) : '';
  const pathToCompare = (pPath: string) => moduleFileName ? path.resolve(pPath, moduleFileName) : pPath;

  const basePath =pathToCompare(base);
  const srcPath =pathToCompare(src);
  const vbePath =pathToCompare(vbe);

  const r = comparePath(basePath, diffTo ==='vbe' ? vbePath : srcPath);

  if (r.differences > 0) {
    outputDiffResult(r, diffTo);
    await outputDiffFile(r, basePath, srcPath, vbePath);
    const ans = await vscode.window.showInformationMessage(confirmMessage, 'Yes', 'No');
    vbeOutput.clear();
    return ans;
  }
  return 'Yes';
}


function outputDiffResult(r: dirCompare.Result, diffTitle: string){
  const result = r.diffSet!.filter(_ => _.state !== 'equal').map(_ => {
    return (_.name1 || _.name2 || '') + ' : ' + _.reason;
  });
  vbeOutput.clear();
  vbeOutput.appendLine(`======== base - ${diffTitle} ===========>`);
  vbeOutput.appendLine(result.join('\n') || '');
}

async function outputDiffFile(r: dirCompare.Result, base: string, src: string, vbe: string){
  const isDir = await common.dirExists(src);

  const dirSrc = isDir ? src : path.dirname(src);
  const dirBase = isDir ? base : path.dirname(base);
  const dirVbe = isDir ? vbe : path.dirname(vbe);

  for (const diff of r.diffSet!){
    if(diff.state !== 'equal'){
      const fileName = (diff.name1 && diff.name2) ? diff.name1 : '';
      const baseFile = path.resolve(dirBase, fileName);
      const vbeFile = path.resolve(dirVbe, fileName);
      await common.fileExists(baseFile) && fse.copyFileSync(baseFile, path.resolve(dirSrc, fileName + '.base'));
      await common.fileExists(vbeFile) && fse.copyFileSync(vbeFile, path.resolve(dirSrc, fileName + '.vbe'));
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
    excludeFilter: '**/*.frx,.base,.current,.git,.gitignore',
    ignoreAllWhiteSpaces: true,
    ignoreEmptyLines: true,
    ignoreCase: true
  };

  const r = dirCompare.compareSync(path1, path2, options);

  return r;
}

export async function deleteModulesInSrc(pathSrc: string, isDiffFile: boolean = false){
  try {

    const isExist = await common.dirExists(pathSrc);
    if (!isExist)
    {
      // no folder
      return;
    }

    const files = fse.readdirSync(pathSrc);

    const extensions = isDiffFile ? ['.bas','.frm','.cls', '.frx', '.base', '.vbe'] : ['.bas','.frm','.cls', '.frx'];
    
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

async function exportModuleAsync(bookPath: string, pathToExport: string, moduleFileName: string) {
  // export modules
  const isSheet =  'True';
  const isForm = 'True';

  const moduleFileNameFix = path.basename(moduleFileName);
  // export
  const { err, status } = runVbs('export.vbs', [bookPath, pathToExport, moduleFileNameFix, isSheet, isForm]);
  if (status !== 0 || err) {
    const error = new Error(err);
    throw (error);
  }

  if (moduleFileNameFix === ''){
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

  if (! await common.fileExists(path.resolve(pathToExport, moduleFileNameFix))){
    throw (Error(`${moduleFileNameFix} is not exported.`)); 
  }
}


async function importModuleSync(pathBook: string, modulePath: string = '') {

  // do import
  const { err, status } = runVbs('import.vbs', [pathBook, modulePath]);
  if (status !== 0 || err) {
    const error = Error(err);
    throw (error);
  }

  //  verify src and exported(once imported)
  //  comparePath does not work with ignore case.
  const { srcDir, baseDir, tempDir } = getVbecmDirs(pathBook, 'importModule');

  // compare dir or file ?
  const moduleFileName = path.basename(modulePath);
  const srcPath = modulePath ? path.resolve(srcDir, moduleFileName) : srcDir;
  const vbePath = modulePath ? path.resolve(tempDir, moduleFileName) : tempDir;

  try {
    await exportModuleAsync(pathBook, tempDir, moduleFileName);

    const r = comparePath(srcPath, vbePath);
    if (r.same === false){
      outputDiffResult(r, 'src, vbe');
      await outputDiffFile(r, baseDir, srcDir, tempDir);
      vscode.window.showWarningMessage('May be Ok, some files are format in VBE engine. Case or Space.');
      //throw(Error('Import verify Error.'));
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

