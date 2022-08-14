// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import path = require('path');
import * as vscode from 'vscode';
import { TextEditor } from 'vscode';
//import { doDiff } from './diff'

import { ConsoleReporter } from '@vscode/test-electron';
import * as fs from 'fs'; 
import * as iconv from 'iconv-lite';


let statusBarVba: vscode.StatusBarItem;

// this method is called when your extension is activated
// your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {
  
  console.log('Congratulations, your extension "vbe-client-mini" is now active!');

  statusBarVba = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Right, 100);
	// statusBarVba.command = myCommandId;
	context.subscriptions.push(statusBarVba);

  // const isUseFormModule = () => vscode.workspace.getConfiguration('vbecm').get<boolean>('useFormModule');
  // const isUseSheetModule = () => vscode.workspace.getConfiguration('vbecm').get<boolean>('useSheetModule');
  
  // export
  const commandExport = vscode.commands.registerCommand(
    'command.export', 
    exportModuleAsync
  );
  context.subscriptions.push(commandExport);

  // import
  const commandImport = vscode.commands.registerCommand(
    'command.import', 
    importModuleAsync
  );
  context.subscriptions.push(commandImport);

  // compile
  const commandCompile = vscode.commands.registerCommand(
    'command.compile', 
    compile
  );
  context.subscriptions.push(commandCompile);

  // run
  const commandRun = vscode.commands.registerTextEditorCommand (
    'editor.run', 
    runAsync
  );
  context.subscriptions.push(commandRun);

  // check out on editor
  const commandCheckOut = vscode.commands.registerCommand(
    'editor.checkout', 
    checkoutAsync
  );
  context.subscriptions.push(commandCheckOut);

  // commit form editor
  const commandCommit = vscode.commands.registerCommand(
    'editor.commit', 
    commitAsync
  );
  context.subscriptions.push(commandCommit);
  
  displayMenu(true);
  // update status bar item once at start
	updateStatusBarItem(false);
}

//End ------------------------------------------------------------------------

// this method is called when your extension is deactivated
export function deactivate() {}


// status bar
function updateStatusBarItem(isVbaWork: boolean): void {
	if (isVbaWork) {
		statusBarVba.text = `[Vba Working]`;
		statusBarVba.show();
	} else {
		statusBarVba.hide();
	}
}

const exportModuleAsync = async (uri:vscode.Uri) => {
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
  const {err, status} = runVbs('export.vbs',[xlsmPath]);
  if (status !== 0 || err){
    showErrorMessage(err);
  } else {
    showInformationMessage("Success export.");
  }
  displayMenu(true);
};

const canExportOrImport = async (srcDir: string, xlsmPath: string): Promise<boolean> =>{
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

const importModuleAsync = async (uri:vscode.Uri) => {
  displayMenu(false);
  //console.log(uri.fsPath);
  const xlsmPath = uri.fsPath;
  const fileDir = path.dirname(xlsmPath);
  const baseName = path.basename(xlsmPath);
  const srcDir = path.resolve(fileDir,'src_' + baseName);

  const isSrcExist = await dirExists(srcDir);
  if (!isSrcExist){
    showErrorMessage(srcDir);
    displayMenu(true);
  }

  const xlsExists = await fileExists(xlsmPath);
  if (!xlsExists){
    showErrorMessage(xlsmPath);
    displayMenu(true);
  }

  if (await canExportOrImport(srcDir,xlsmPath)){
    // go forward
  }else{
    displayMenu(true);
    return;
  }

  const modulePath = '';
  const {err, status} = runVbs('import.vbs', [xlsmPath, modulePath]);

  if (status !== 0 || err){
    showErrorMessage(err);
  } else {
    showInformationMessage("Success import.");
  }
  displayMenu(true);
};

const compile = (uri:vscode.Uri) => {
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

const runAsync = async (textEditor: TextEditor, edit: vscode.TextEditorEdit, uri:vscode.Uri) => {
  displayMenu(false);
  // excel file path
  const xlsmPath = await getExcelPathFromModule(uri);

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

const commitAsync = async (uri:vscode.Uri) => {
  displayMenu(false);
  const importModulesVbs = path.resolve(getVbsPath(),'import.vbs');
  const xlsmPath = await getExcelPathFromModule(uri);
  const modulePath = uri.fsPath;
  
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
const checkoutAsync = async (uri:vscode.Uri) => {
  displayMenu(false);

  const xlsmPath = await getExcelPathFromModule(uri);
  const modulePath = uri.fsPath;

  const fileDir = path.dirname(xlsmPath);
  const baseName = path.basename(xlsmPath);
  
  // test really checkout (export form excel)
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
 * set display vbe menu on or off
 * @param isOn
 */
const displayMenu = (isOn : boolean) =>{
  vscode.commands.executeCommand('setContext', 'vbecm.showVbsCommand', isOn);
  updateStatusBarItem(!isOn);
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