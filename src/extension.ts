// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import path = require('path');
import * as vscode from 'vscode';
import { TextEditor } from 'vscode';
//import { doDiff } from './diff'

import { ConsoleReporter } from '@vscode/test-electron';
import * as fs from 'fs'; 
import * as iconv from 'iconv-lite';

// this method is called when your extension is activated
// your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {
  
  console.log('Congratulations, your extension "vbe-client-mini" is now active!');

  displayMenu(true);

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

  // commit form editor
  const commandCommit = vscode.commands.registerCommand(
    'editor.commit', 
    commitAsync
  );
}

// this method is called when your extension is deactivated
export function deactivate() {}

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
  }

  // test number of modules
  // const moduleCountSrc = getSrcModules(srcDir);
  // const moduleCountVbe = getVbeModules(xlsmPath);
  // if (moduleCountSrc !== moduleCountVbe){
  //   const ans = await vscode.window.showInformationMessage(`Modules is not same between vbe(${moduleCountSrc}) and src(${moduleCountVbe}). Do you force?`, "Yes", "No");
  //   if (ans === 'No'){
  //     displayMenu(true);
  //     return;
  //   }
  // }
  if (await testModulesAndForce(srcDir,xlsmPath)){
    //
  }else{
    return;
  }
  
  // export
  const {err, status} = runVbs('export.vbs',[xlsmPath]);
  if (status !== 0 || err){
    vscode.window.showErrorMessage(err);
  } else {
    showInformationMessage("Success export.");
  }
  displayMenu(true);
};


const testModulesAndForce = async (srcDir: string, xlsmPath: string): Promise<boolean> =>{
  const moduleCountSrc = getSrcModules(srcDir);
  const moduleCountVbe = getVbeModules(xlsmPath);
  if (moduleCountSrc !== moduleCountVbe){
    const ans = await vscode.window.showInformationMessage(`Modules is not same between vbe(${moduleCountVbe}) and src(${moduleCountSrc}). Do you force?`, "Yes", "No");
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

  if (await testModulesAndForce(srcDir,xlsmPath)){
    //
  }else{
    return;
  }

  const modulePath = '';
  const {err, status} = runVbs('import.vbs', [xlsmPath, modulePath]);

  if (status !== 0 || err){
    vscode.window.showErrorMessage(err);
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
    vscode.window.showErrorMessage(err);
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
      vscode.window.showErrorMessage(err);
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
    vscode.window.showErrorMessage(err);
  } else {
    showInformationMessage("Success commit.");
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

/**
 * 
 * @param param vbs, ...params
 * @returns messages
 */
const runVbs = (script: string, param:string[]) =>{
  //console.log(param);
  const scriptPath = path.resolve(getVbsPath(), script);
  const spawn = require('child_process').spawnSync,
  vbs = spawn('cscript.exe', ['//Nologo', scriptPath, ...param] );
  const err = s2u(vbs.stderr);
  const out = s2u(vbs.stdout);
  console.log( `stderr: ${err}` );
  console.log( `stdout: ${out}` );
  console.log( `status: ${vbs.status}` );
  return {err, out,  status: vbs.status};
};

// shift jis 2 utf8
// only japanese
const s2u = (sb: Buffer) =>{
  const vbsEncode = vscode.workspace.getConfiguration('vbecm').get<string>('vbsEncode') || 'windows-31j';
  return iconv.decode(sb, vbsEncode);
};

const getSrcModules = (dirPath: string) => {
  const allDirents = fs.readdirSync(dirPath, { withFileTypes: true });
  const count = allDirents.filter((file) => path.extname(file.name).match(/^\.cls$|^\.bas$|^\.frm$/) && file.isFile).length;
  return count;
};

const getVbeModules = (xlsmPath : string) : number =>{
  const a= runVbs('getModules.vbs',[xlsmPath]);
  if (a.status !== 0 || a.err){
    return -1;
  } else {
    return parseInt(a.out, 10);
  }
};

const getVbsPath = () =>{
  const rootFolder = path.dirname(path.dirname(__filename));
  const vbsPath = path.resolve(rootFolder, "vbs");
  return vbsPath;
};