import * as path from 'path';
import * as vscode from 'vscode';
import { TextEditor } from 'vscode';
import * as fs from 'fs';
import * as fse from 'fs-extra';
import * as iconv from 'iconv-lite';
import * as stBar from './statusBar';
import { vbeOutput } from './vbeOutput';
import { v4 as uuidv4 } from 'uuid';
import {spawnSync} from 'child_process';
import { compareSync, Options } from 'dir-compare';
import { dirExists, fileExists } from './common';
import * as common from './common';

/**
 * get vbecmDirs
 * @param uriBook
 * @returns 
 */
export function vbecmDirs(xlsmPath: string) {
  const fileDir = path.dirname(xlsmPath);
  const baseName = path.basename(xlsmPath);
  const srcDir = path.resolve(fileDir, 'src_' + baseName);
  const baseDir = path.resolve(srcDir, '.base');
  const tempDir = path.resolve(fileDir, 'src_' + uuidv4());

  return {
    xlsmPath,
    srcDir,
    baseDir,
    tempDir
  };
}


export type TestConfirm = { (dir1: string, dir2: string, diffTitle: string): Promise<boolean>; };



export async function exportModules(pathBook: string, testConfirm?:TestConfirm )  {
  const {srcDir, baseDir, tempDir} = vbecmDirs(pathBook);
    
  // export to uuid folder

  try {
    exportModuleSync(pathBook, tempDir, '');  
  } catch (error) {
    console.log(error);
    throw(error);
  }
 

  // test already exported
  const isExistSrc = await dirExists(srcDir);

  // check tempDir and srcDir
  if (isExistSrc && testConfirm){
    const ans = await testConfirm(tempDir, srcDir, 'Diff between xlsm(a) and src.');
    if (ans === false){
      await fs.promises.rm(tempDir, { recursive: true, force: true });
      return;
    }
  }

  // copy temp to real folder
  try {
    deleteModulesInSrc(srcDir);
    await fse.copy(tempDir, srcDir);
    deleteModulesInSrc(baseDir);
    await fse.copy(tempDir, baseDir);
  }
  catch (e) {
    throw(e);
  }
  finally{
    if (await dirExists(srcDir)){
      await fs.promises.rm(tempDir, { recursive: true, force: true }); 
    }
  }
};

export async function importModules(pathBook: string, testConfirm?:TestConfirm) {

  const { xlsmPath, srcDir, baseDir, tempDir } = vbecmDirs(pathBook);

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
    exportModuleSync(pathBook, tempDir, '');
  }
  catch(e)
  {
    throw(e);
  }

  // check tempDir and srcDir
  const undefinedTestFunc = !testConfirm;
  const isGoForward = undefinedTestFunc || await testConfirm(tempDir, baseDir, '');
  if (isGoForward === false){
    await fs.promises.rm(tempDir, { recursive: true, force: true });
    return;
  }

  // import main
  try {
    importModuleSync(pathBook);
  }
  catch (e) {
    throw(e);
  }
  finally{
    await fs.promises.rm(tempDir, { recursive: true, force: true });
  }

  // update src folder(copy temp to real folder)
  try{
    await exportModules(pathBook);
  }
  catch(e)
  {
    throw(e);
  }
}

export async function compareFoldersAndGo(path1: string, path2: string, diffTitle: string, confirmMessage: string) {
  // compare options
  const options: Options = {
    compareSize: true,
    compareContent: true,
    skipSubdirs: true,
    excludeFilter: '**/*.frx,.base,.current,.git,.gitignore'
  };

  const r = compareSync(path1, path2, options);
  if (r.differences > 0) {
    const result = r.diffSet!.filter(_ => _.state !== 'equal').map(_ => {
      return (_.name1 || _.name2 || '') + ' : ' + _.reason;
    });
    vbeOutput.clear();
    vbeOutput.appendLine(`======== ${diffTitle} ===========>`);
    vbeOutput.appendLine(result.join('\n') || '');

    const ans = await vscode.window.showInformationMessage(confirmMessage, 'Yes', 'No');

    return ans;
  }
  return 'Yes';
}


function exportModuleSync(xlsmPath:string, pathToExport: string, moduleName: string) {
  // export
  const { err, status } = runVbs('export.vbs', [xlsmPath, pathToExport, moduleName]);
  if (status !== 0 || err) {
    const error = new Error(err);
    throw (error);
  }
}

export async function exportFrxModules(pathBook: string){

  if (!fileExists(pathBook)){
    throw(Error(`Excel file does not exist.: ${pathBook}`));
  }

  const {srcDir, baseDir, tempDir} = vbecmDirs(pathBook);

  // export to uuid folder
  try{
    exportModuleSync(pathBook, tempDir, '');
  }
  catch(e)
  {
    throw(e);
  }

  const fileNameList = fs.readdirSync(tempDir);
  const targetFileNames = fileNameList.filter(RegExp.prototype.test, /.*\.frx$/i);

  targetFileNames.forEach(fileName => {
    fs.copyFileSync(path.resolve(tempDir,fileName),path.resolve(srcDir,fileName));
    fs.copyFileSync(path.resolve(tempDir,fileName),path.resolve(baseDir,fileName));
  });

  await fs.promises.rm(tempDir, { recursive: true, force: true });
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

  const { err, status } = runVbs('runVba.vbs', [xlsmPath, moduleName + '.' +funcName]);
  if (status !== 0 || err) {
    throw(Error(err));
  }
}


export function compile(bookPath: string) {
  const { err, status } = runVbs('compile.vbs', [bookPath]);
  if (status !== 0 || err) {
    throw(Error(err));
  }
}


export function closeBook(bookPath: string){
  // get path, but close a same name book.
  const { err, status } = runVbs('closeBook.vbs', [bookPath]);
  if (status !== 0 || err) {
    throw(Error(err));
  }

}


export async function commitModule(xlsmPath: string, modulePath:string, testConfirm?:TestConfirm) {
  if (!fileExists(xlsmPath)) {
    throw(Error(`Excel file does not exist to commit.: ${modulePath}`));
  }

  //module in excel modified?
  const {srcDir, baseDir, tempDir} = vbecmDirs(xlsmPath);
  exportModuleSync(xlsmPath, tempDir, '');

  // check tempDir and srcDir
  const moduleBase = path.basename(modulePath);

  const doTest = testConfirm !== undefined;
  const goNext = !doTest || await testConfirm(path.resolve(tempDir, moduleBase), path.resolve(baseDir, moduleBase), '');
  if (goNext === false){
    await fs.promises.rm(tempDir, { recursive: true, force: true });
    return;
  }

  const { err, status } = runVbs('import.vbs', [xlsmPath, modulePath]);

  if (status !== 0) {
    throw(Error(err));
  }
  
  // export commit files
  exportModuleSync(xlsmPath, srcDir, path.resolve(srcDir, moduleBase));
  await fse.copy(path.resolve(srcDir, moduleBase), path.resolve(baseDir, moduleBase));
  await fs.promises.rm(tempDir, { recursive: true, force: true });
}

export async function updateModule(xlsmPath: string, modulePath:string, testConfirm? :TestConfirm) {
  // old version, checkout

  // if the file does not exist, return ''
  if (!fileExists(xlsmPath)) {
    throw(Error(`Excel file does not exist to update src.: ${modulePath}`));
  }

  // test really checkout (export form excel)
  // test always
  const goNext = testConfirm === undefined;
  const ans = goNext || await vscode.window.showInformationMessage('Do you want to checkout and overwrite?', 'Yes', 'No');
  if (ans === 'No') {
    return;
  }

  //module in excel modified?
  const {srcDir, baseDir, tempDir} = vbecmDirs(xlsmPath);

  const moduleBase = path.basename(modulePath);

  try{
    exportModuleSync(xlsmPath, srcDir, path.resolve(srcDir, moduleBase));
    await fse.copy(path.resolve(srcDir, moduleBase), path.resolve(baseDir, moduleBase));
  }
  catch(e)
  {
    throw(e);
  }

  await fs.promises.rm(tempDir, { recursive: true, force: true });
}


function importModuleSync(xlsmPath: string) {
  const modulePath = '';
  const { err, status } = runVbs('import.vbs', [xlsmPath, modulePath]);

  if (status !== 0 || err) {
    const error = Error(err);
    throw (error);
  }
}


type RunVbs = {err: string, out: string, retValue: string, status: number | null};
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

export function deleteModulesInSrc(pathSrc: string){
  try {

    if (!common.dirExists(pathSrc))
    {
      // フォルダが存在しないので 消さない
      return;
    }

    //ファイルとディレクトリのリストが格納される(配列)
    const files = fs.readdirSync(pathSrc);
    
    //ディレクトリのリストに絞る
    const moduleList = files.forEach((file) => {
        const fullPath = path.resolve(pathSrc, file);
        const isFile = fs.statSync(fullPath).isFile();
        const ext  = path.extname(file).toLowerCase();
        if (isFile && ['.bas','.frm','.cls', '.frx'].includes(ext)){
          fs.rmSync(fullPath);
        }
    });   
  } catch (error) {
    throw(error);
  }
}