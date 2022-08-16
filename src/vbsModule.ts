import * as path from 'path';
import * as vscode from 'vscode';
import { TextEditor } from 'vscode';
import * as fs from 'fs'; 
import * as iconv from 'iconv-lite';
import * as stBar from './statusBar';


export const exportModuleAsync = async (uri:vscode.Uri) => {
  displayMenu(false);

  const xlsmPath = uri.fsPath;
  const fileDir = path.dirname(xlsmPath);
  const baseName = path.basename(xlsmPath);
  const srcDir = path.resolve(fileDir,'src_' + baseName);
  
  // test already exported
  const isExist = await dirExists(srcDir);
  if (isExist){
    const ans = await vscode.window.showInformationMessage("Already exported. Do you want to export?", "Yes", "No");
    if (ans === 'No'){
      displayMenu(true);
      return;
    }

    // test module count diff
    if (await canExportOrImport(srcDir,xlsmPath)){
      // go forward
    }else{
      displayMenu(true);
      return;
    }
  }

  // export
  exportModuleSync(uri);
};

export const exportModuleSync = (uri:vscode.Uri) => {
  displayMenu(false);

  const xlsmPath = uri.fsPath;
  
  // export
  const {err, status} = runVbs('export.vbs',[xlsmPath]);
  if (status !== 0 || err){
    showErrorMessage(err);
  } else {
    showInformationMessage("Success export.");
  }
  displayMenu(true);
};

export const canExportOrImport = async (srcDir: string, xlsmPath: string): Promise<boolean> =>{
  const moduleCountSrc = getCountSrcModules(srcDir);
  const moduleCountVbe = getCountVbeModules(xlsmPath);
  if (moduleCountSrc !== moduleCountVbe){
    const ans = await vscode.window.showInformationMessage(
      `Modules is not same between vbe(${moduleCountVbe}) and src(${moduleCountSrc}). Do you force?`, "Yes", "No");
    if (ans === 'No'){
      displayMenu(true);
      return false;
    }
  }
  return true;
};

/**
 * 
 * @param uri excel path
 * @returns 
 */
 export const importModuleAsync = async (uri:vscode.Uri) => {
  displayMenu(false);
  //console.log(uri.fsPath);
  const xlsmPath = uri.fsPath;
  const fileDir = path.dirname(xlsmPath);
  const baseName = path.basename(xlsmPath);
  const srcDir = path.resolve(fileDir,'src_' + baseName);

  const isSrcExist = await dirExists(srcDir);
  if (!isSrcExist){
    showErrorMessage(`Source folder does not exist: ${srcDir}`);
    displayMenu(true);
  }

  const xlsExists = await fileExists(xlsmPath);
  if (!xlsExists){
    showErrorMessage(`Excel file does not exist: ${xlsmPath}`);
    displayMenu(true);
  }

  if (await canExportOrImport(srcDir,xlsmPath)){
    // go forward
  }else{
    displayMenu(true);
    return;
  }

  importModuleSync(uri);
};


export const importModuleSync = (uri:vscode.Uri) => {
  displayMenu(false);
  const xlsmPath = uri.fsPath;
  const modulePath = '';
  const {err, status} = runVbs('import.vbs', [xlsmPath, modulePath]);

  if (status !== 0 || err){
    showErrorMessage(err);
  } else {
    showInformationMessage("Success import.");
  }
  displayMenu(true);
};



export const compile = (uri:vscode.Uri) => {
  displayMenu(false);
  console.log(uri.fsPath);
  const compileVbs = path.resolve(getVbsPath(), 'compile.vbs');
  const xlsmPath = uri.fsPath;

  const {err, status} = runVbs('compile.vbs', [xlsmPath]);

  if (status !== 0 || err){
    showErrorMessage(err);
  } else {
    showInformationMessage("Success compile.");
  }
  displayMenu(true);
};

export const runAsync = async (textEditor: TextEditor, edit: vscode.TextEditorEdit, uri:vscode.Uri) => {
  displayMenu(false);
  // excel file path
  const xlsmPath = await getExcelPathFromModule(uri);
  const modulePath = uri.fsPath;
  if (xlsmPath === ''){
    showErrorMessage(`Excel file does not exist to run.: ${modulePath}`);
    displayMenu(true);
    return;
  }

  // sub function
  const activeLine = textEditor.document.lineAt(textEditor.selection.active.line).text;
  const vbaSub = activeLine.match(/Sub (.*)\(\s*\)/i);
  const funcName = vbaSub !== null && vbaSub.length === 2 && vbaSub[1];

  if (funcName){
    const {err, status} = runVbs('runVba.vbs', [xlsmPath, funcName] );

    if (status !== 0 || err){
      showErrorMessage(err);
    } else {
      showInformationMessage("Success run.");
    }
  }
  else{
    console.log('No sub function is selected.');
  }
  displayMenu(true);
};

export const commitAsync = async (uri:vscode.Uri) => {
  displayMenu(false);

  const xlsmPath = await getExcelPathFromModule(uri);
  const modulePath = uri.fsPath;
  if (xlsmPath === ''){
    showErrorMessage(`Excel file does not exist to commit.: ${modulePath}`);
    displayMenu(true);
    return;
  }

  const {err, status} = runVbs('import.vbs',[xlsmPath, modulePath] );

  if (status !== 0){
    showErrorMessage(err);
  } else {
    showInformationMessage("Success commit.");
  }
  displayMenu(true);
};


/**
 * 
 * @param uri module path
 * @returns 
 */
 export const checkoutAsync = async (uri:vscode.Uri) => {
  displayMenu(false);
  
  // if the file does not exist, return ''
  const xlsmPath = await getExcelPathFromModule(uri);
  const modulePath = uri.fsPath;
  if (xlsmPath === ''){
    showErrorMessage(`Excel file does not exist to checkout.: ${modulePath}`);
    displayMenu(true);
    return;
  }
  
  // test really checkout (export form excel)
  // test always
  const ans = await vscode.window.showInformationMessage("Do you want to checkout and overwrite?", "Yes", "No");
  if (ans === 'No'){
    displayMenu(true);
    return;
  }

  // export
  const {err, status} = runVbs('export.vbs',[xlsmPath, '0', modulePath]);
  if (status !== 0 || err){
    showErrorMessage(err);
  } else {
    showInformationMessage("Success checkout.");
  }
  displayMenu(true);
};

/**
 * set display vbe menu on or off
 * @param isOn
 */
 const displayMenu = (isOn : boolean) =>{
  vscode.commands.executeCommand('setContext', 'vbecm.showVbsCommand', isOn);
  stBar.updateStatusBarItem(!isOn);
};


/**
 * 
 * @param message 
 */
const showInformationMessage =(message: string) =>{
  const date = new Date();
  const dateString = date.getHours().toString().padStart(2,"0") + ":" + date.getMinutes().toString().padStart(2,"0");
  vscode.window.showInformationMessage(dateString + " , " + message);
};


const showErrorMessage =(message: string) =>{
  const date = new Date();
  const dateString = date.getHours().toString().padStart(2,"0") + ":" + date.getMinutes().toString().padStart(2,"0");
  vscode.window.showErrorMessage(dateString + " , " + message);
};


/**
 * get Excel path
 * @param uri 
 * @returns 
 */
const getExcelPathFromModule = async (uri:vscode.Uri)  => {
  const dirParent = path.dirname(uri.fsPath);
  const dirForBook = path.dirname(dirParent);
  const bookFileName = path.basename(path.dirname(uri.fsPath)).slice("src_".length);
  const xlsPath = path.resolve(dirForBook, bookFileName);
  const isFile = await fileExists(xlsPath);
  return isFile ? xlsPath : '';
};




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

import {spawnSync} from 'child_process';

type RunVbs = {err: string, out: string, retValue: string, status: number | null};
/**
 * 
 * @param param vbs, ...params
 * @returns messages
 */
const runVbs = (script: string, param:string[]) =>{
  //console.log(param);
  try{
    const scriptPath = path.resolve(getVbsPath(), script);
    const vbs = spawnSync('cscript.exe', ['//Nologo', scriptPath, ...param] );
    const err = s2u(vbs.stderr);
    const out = s2u(vbs.stdout);
    const retValue = out.split('\r\n').slice(-2)[0];
    console.log( `>====================vbs run in=======================` );
    console.log( `Script: ${scriptPath}` );
    console.log( `stderr: ${err}` );
    console.log( `stdout: ${out}` );
    console.log( `retVal: ${retValue}` );
    console.log( `status: ${vbs.status}` );
    console.log( `====================vbs run out=======================` );
    return {err, out, retValue, status: vbs.status} ;
  }
  catch (e : unknown){
    const errMessage = (e instanceof Error) ? e.message : 'vbs run error.';
    return {err: errMessage, out: '', retValue : '', status: 10};
  }
};

// shift jis 2 utf8
// only japanese
const s2u = (sb: Buffer) =>{
  const vbsEncode = vscode.workspace.getConfiguration('vbecm').get<string>('vbsEncode') || 'windows-31j';
  return iconv.decode(sb, vbsEncode);
};

const getCountSrcModules = (dirPath: string) => {
  const files = fs.readdirSync(dirPath, { withFileTypes: true });
  const count = files.filter((file) => 
    path.extname(file.name).match(/^\.cls$|^\.bas$|^\.frm$/) && file.isFile).length;
  return count;
};

const getCountVbeModules = (xlsmPath : string) : number =>{
  const retModuleCount = runVbs('getModules.vbs',[xlsmPath]);
  if (retModuleCount.status !== 0 || retModuleCount.err){
    return -1;
  } else {
    return parseInt(retModuleCount.retValue, 10);
  }
};

const getVbsPath = () =>{
  const rootFolder = path.dirname(path.dirname(__filename));
  const vbsPath = path.resolve(rootFolder, "vbs");
  return vbsPath;
};