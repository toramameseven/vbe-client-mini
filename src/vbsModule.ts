import * as path from 'path';
import * as vscode from 'vscode';
import { TextEditor } from 'vscode';
import * as fs from 'fs';
import * as fse from 'fs-extra';
import * as iconv from 'iconv-lite';
import * as stBar from './statusBar';
import { myOutput } from './myOutput';
import { v4 as uuidv4 } from 'uuid';
import {spawnSync} from 'child_process';
import { compareSync, Options } from 'dir-compare';


function dirs(uri: vscode.Uri) {
  const xlsmPath = uri.fsPath;
  const fileDir = path.dirname(xlsmPath);
  const baseName = path.basename(xlsmPath);
  const srcDir = path.resolve(fileDir, 'src_' + baseName);
  const baseDir = path.resolve(srcDir, '.base');
  const uuid = uuidv4();
  const tempDir = path.resolve(fileDir, 'src_' + uuid);

  return {
    xlsmPath,
    srcDir,
    baseDir,
    tempDir
  };
}


export async function exportModuleAsync(uri:vscode.Uri)  {
  displayMenu(false);

  const {srcDir, baseDir, tempDir} = dirs(uri);
  
  // test already exported
  const isExistSrc = await dirExists(srcDir);

  // export to uuid folder
  exportModuleSync(uri, tempDir, '');

  // check tempDir and srcDir
  if (isExistSrc){
    const ans = await compareFoldersAndGo(tempDir, srcDir, 'Differ excel form src. Check Output Window. Do you want to export?');
    if (ans !== 'Yes'){
      await fs.promises.rm(tempDir, { recursive: true, force: true });
      displayMenu(true);
      return;
    }
  }

  await updateSrcFolder({ uriBook: uri, tempDir });
};


async function updateSrcFolder({ uriBook, tempDir }: { uriBook: vscode.Uri; tempDir: string; }) {

  const { srcDir, baseDir } = dirs(uriBook);
  // copy temp to real folder
  try {
    const isExistSrc = await dirExists(srcDir);
    if (isExistSrc) {
      await fs.promises.rm(srcDir, { recursive: true, force: true });
    }

    await fse.copy(tempDir, srcDir);
    await fse.copy(tempDir, baseDir);

    await fs.promises.rm(tempDir, { recursive: true, force: true });
    showInformationMessage("Success export.");
  }
  catch (e) {
    if (e instanceof Error) {
      showErrorMessage(e.message);
    }
    console.log(e);
  }
  finally {
    displayMenu(true);
  }
}

export async function updateFrxModule(uriBook:vscode.Uri){

  const {srcDir, baseDir, tempDir} = dirs(uriBook);
  // export all
  exportModuleSync(uriBook, tempDir, '');

  const fileNameList = fs.readdirSync(tempDir);
  const targetFileNames = fileNameList.filter(RegExp.prototype.test, /.*\.frx$/i);
  // console.log(targetFileNames);
  
  targetFileNames.forEach(fileName => {
    fs.copyFileSync(path.resolve(tempDir,fileName),path.resolve(srcDir,fileName));
    fs.copyFileSync(path.resolve(tempDir,fileName),path.resolve(baseDir,fileName));
  });

  await fs.promises.rm(tempDir, { recursive: true, force: true });
}



async function compareFoldersAndGo(path1: string, path2: string, confirmMessage: string) {
  const options: Options = {
    compareSize: true,
    compareContent: true,
    skipSubdirs: true,
    excludeFilter: "**/*.frx,.base,.current"
  };
  const r = compareSync(path1, path2, options);
  if (r.differences > 0) {
    const result = r.diffSet!.filter(_ => _.state !== 'equal').map(_ => {
      return (_.name1 || _.name2 || '') + ' : ' + _.reason;
    });
    myOutput.clear();
    myOutput.appendLine('===================>');
    myOutput.appendLine(result.join('\n') || '');

    const ans = await vscode.window.showInformationMessage(confirmMessage, "Yes", "No");

    return ans;
  }
  return 'Yes';
}


export function exportModuleSync(uriBook: vscode.Uri | string, pathToExport: string, moduleName: string) {
  const xlsmPath = typeof uriBook === "string" ? uriBook : uriBook.fsPath;
  // export
  const { err, status } = runVbs('export.vbs', [xlsmPath, pathToExport, moduleName]);
  if (status !== 0 || err) {
    const error = Error(err);
    throw (error);
  }
}


export async function canExportOrImport(srcDir: string, xlsmPath: string): Promise<boolean> {
  const moduleCountSrc = getCountSrcModules(srcDir);
  const moduleCountVbe = getCountVbeModules(xlsmPath);

  if (moduleCountSrc !== moduleCountVbe) {
    const ans = await vscode.window.showInformationMessage(
      `Modules is not same between vbe(${moduleCountVbe}) and src(${moduleCountSrc}). Do you force?`, "Yes", "No");
    myOutput.appendLine(ans?.toString() || '');
    if (ans === 'No') {
      displayMenu(true);
      return false;
    }
  }
  return true;
}

/**
 * 
 * @param uri excel path
 * @returns 
 */
 export async function importModuleAsync(uri: vscode.Uri) {
  displayMenu(false);

  const { xlsmPath, srcDir, baseDir, tempDir } = dirs(uri);

  const isSrcExist = await dirExists(srcDir);
  if (!isSrcExist) {
    showErrorMessage(`Source folder does not exist: ${srcDir}`);
    displayMenu(true);
  }

  const xlsExists = await fileExists(xlsmPath);
  if (!xlsExists) {
    showErrorMessage(`Excel file does not exist: ${xlsmPath}`);
    displayMenu(true);
  }

  //module in excel modified?
  exportModuleSync(uri, tempDir, '');

  // check tempDir and srcDir
  const ans = await compareFoldersAndGo(tempDir, baseDir, 'Excel modules are modified form export. Check Output Window. Do you want to import?');
  if (ans !== 'Yes') {
    await fs.promises.rm(tempDir, { recursive: true, force: true });
    displayMenu(true);
    return;
  }


  // import main
  try {
    importModuleSync(uri);
  }
  catch (e) {
    if (e instanceof Error) {
      showErrorMessage(e.message);
    }
    console.log(e);
  }

  // update src folder(copy temp to real folder)
  await updateSrcFolder({ uriBook: uri, tempDir });
}


export function importModuleSync(uri: vscode.Uri) {
  displayMenu(false);
  const xlsmPath = uri.fsPath;
  const modulePath = '';
  const { err, status } = runVbs('import.vbs', [xlsmPath, modulePath]);

  if (status !== 0 || err) {
    const error = Error(err);
    throw (error);
  }
}



export function compile(uri: vscode.Uri) {
  displayMenu(false);
  console.log(uri.fsPath);
  const compileVbs = path.resolve(getVbsPath(), 'compile.vbs');
  const xlsmPath = uri.fsPath;

  const { err, status } = runVbs('compile.vbs', [xlsmPath]);

  if (status !== 0 || err) {
    showErrorMessage(err);
  } else {
    showInformationMessage("Success compile.");
  }
  displayMenu(true);
}

export async function runAsync(textEditor: TextEditor, edit: vscode.TextEditorEdit, uri: vscode.Uri) {
  displayMenu(false);
  // excel file path
  const xlsmPath = await getExcelPathFromModule(uri);
  const modulePath = uri.fsPath;
  if (xlsmPath === '') {
    showErrorMessage(`Excel file does not exist to run.: ${modulePath}`);
    displayMenu(true);
    return;
  }

  // sub function
  const activeLine = textEditor.document.lineAt(textEditor.selection.active.line).text;
  const vbaSub = activeLine.match(/Sub (.*)\(\s*\)/i);
  const funcName = vbaSub !== null && vbaSub.length === 2 && vbaSub[1];

  if (funcName) {
    const { err, status } = runVbs('runVba.vbs', [xlsmPath, funcName]);

    if (status !== 0 || err) {
      showErrorMessage(err);
    } else {
      showInformationMessage("Success run.");
    }
  }
  else {
    console.log('No sub function is selected.');
  }
  displayMenu(true);
}

export async function commitAsync(uriModule: vscode.Uri) {
  displayMenu(false);

  const xlsmPath = await getExcelPathFromModule(uriModule);
  const modulePath = uriModule.fsPath;
  if (xlsmPath === '') {
    showErrorMessage(`Excel file does not exist to commit.: ${modulePath}`);
    displayMenu(true);
    return;
  }

  //module in excel modified?
  const {srcDir, baseDir, tempDir} = dirs(vscode.Uri.file(xlsmPath));
  exportModuleSync(vscode.Uri.file(xlsmPath), tempDir, '');

  // check tempDir and srcDir
  const moduleBase = path.basename(modulePath);
  const ans = await compareFoldersAndGo(
    path.resolve(tempDir, moduleBase), 
    path.resolve(baseDir, moduleBase), 
    `Excel module ${moduleBase} are modified form export. Check Output Window. Do you want to import?`);

  if (ans !== 'Yes') {
    await fs.promises.rm(tempDir, { recursive: true, force: true });
    displayMenu(true);
    return;
  }

  const { err, status } = runVbs('import.vbs', [xlsmPath, modulePath]);

  if (status !== 0) {
    showErrorMessage(err);
  } else {
    showInformationMessage("Success commit.");
  }


  // export commit files
  exportModuleSync(xlsmPath, srcDir, path.resolve(srcDir, moduleBase));
  exportModuleSync(xlsmPath, baseDir, path.resolve(baseDir, moduleBase));
  await fs.promises.rm(tempDir, { recursive: true, force: true });

  displayMenu(true);
}


/**
 * 
 * @param uri module path
 * @returns 
 */
 export async function checkoutAsync(uriModule: vscode.Uri) {
  displayMenu(false);

  // if the file does not exist, return ''
  const xlsmPath = await getExcelPathFromModule(uriModule);
  const modulePath = uriModule.fsPath;
  if (xlsmPath === '') {
    showErrorMessage(`Excel file does not exist to checkout.: ${modulePath}`);
    displayMenu(true);
    return;
  }

  // test really checkout (export form excel)
  // test always
  const ans = await vscode.window.showInformationMessage("Do you want to checkout and overwrite?", "Yes", "No");
  if (ans === 'No') {
    displayMenu(true);
    return;
  }

  //module in excel modified?
  const {srcDir, baseDir, tempDir} = dirs(vscode.Uri.file(xlsmPath));

  const moduleBase = path.basename(modulePath);

  try{
    exportModuleSync(xlsmPath, srcDir, path.resolve(srcDir, moduleBase));
    exportModuleSync(xlsmPath, baseDir, path.resolve(baseDir, moduleBase));
    showInformationMessage("Success checkout.");
  }
  catch(e)
  {
    if (e instanceof Error) {
      showErrorMessage(e.message);
    }
    console.log(e);
  }

  await fs.promises.rm(tempDir, { recursive: true, force: true });
  displayMenu(true);
}

/**
 * set display vbe menu on or off
 * @param isOn
 */
function displayMenu(isOn: boolean) {
  vscode.commands.executeCommand('setContext', 'vbecm.showVbsCommand', isOn);
  stBar.updateStatusBarItem(!isOn);
}

function showInformationMessage(message: string) {
  vscode.window.showInformationMessage(getHourMinute() + " , " + message);
}

function showErrorMessage(message: string) {
  vscode.window.showErrorMessage(getHourMinute() + " , " + message);
}

function getHourMinute(){
  const date = new Date();
  const dateString = date.getHours().toString().padStart(2, "0") + ":" + date.getMinutes().toString().padStart(2, "0");
  return dateString;
}


/**
 * get Excel path
 * @param uri 
 * @returns 
 */
async function getExcelPathFromModule(uri: vscode.Uri) {
  const dirParent = path.dirname(uri.fsPath);
  const dirForBook = path.dirname(dirParent);
  const bookFileName = path.basename(path.dirname(uri.fsPath)).slice("src_".length);
  const xlsPath = path.resolve(dirForBook, bookFileName);
  const isFile = await fileExists(xlsPath);
  return isFile ? xlsPath : '';
}


/**
 * test file exits
 * @param filepath
 * @returns 
 */
async function fileExists(filepath: string) {
  try {
    const res = (await fs.promises.lstat(filepath)).isFile();
    return (res);
  } catch (e) {
    return false;
  }
}

/**
 * test dir exists
 * @param filepath
 * @returns 
 */
async function dirExists(filepath: string) {
  try {
    const res = (await fs.promises.lstat(filepath)).isDirectory();
    return (res);
  } catch (e) {
    return false;
  }
}



type RunVbs = {err: string, out: string, retValue: string, status: number | null};
/**
 * 
 * @param param vbs, ...params
 * @returns messages
 */
function runVbs(script: string, param: string[]) {
  //console.log(param);
  try {
    const scriptPath = path.resolve(getVbsPath(), script);
    const vbs = spawnSync('cscript.exe', ['//Nologo', scriptPath, ...param]);
    const err = s2u(vbs.stderr);
    const out = s2u(vbs.stdout);
    const retValue = out.split('\r\n').slice(-2)[0];
    console.log(`>====================vbs run in=======================`);
    console.log(`Script: ${scriptPath}`);
    console.log(`stderr: ${err}`);
    console.log(`stdout: ${out}`);
    console.log(`retVal: ${retValue}`);
    console.log(`status: ${vbs.status}`);
    console.log(`====================vbs run out=======================`);
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

function getCountSrcModules(dirPath: string){
  const files = fs.readdirSync(dirPath, { withFileTypes: true });
  const count = files.filter((file) => 
    path.extname(file.name).match(/^\.cls$|^\.bas$|^\.frm$/) && file.isFile).length;
  return count;
};

function getCountVbeModules(xlsmPath : string) : number {
  const retModuleCount = runVbs('getModules.vbs',[xlsmPath]);
  if (retModuleCount.status !== 0 || retModuleCount.err){
    return -1;
  } else {
    return parseInt(retModuleCount.retValue, 10);
  }
}

function getVbsPath(){
  const rootFolder = path.dirname(path.dirname(__filename));
  const vbsPath = path.resolve(rootFolder, "vbs");
  return vbsPath;
}