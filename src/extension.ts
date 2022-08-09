// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import path = require('path');
import * as vscode from 'vscode';
import { TextEditor } from 'vscode';
//import { doDiff } from './diff'

import { ConsoleReporter } from '@vscode/test-electron';

// this method is called when your extension is activated
// your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {
  
  console.log('Congratulations, your extension "vbe-client-mini" is now active!');

  // get vbs module path
  const rootFolder = path.dirname(path.dirname(__filename));
  const vbsPath = path.resolve(rootFolder, "vbs");


  const isUseFormModule = () => vscode.workspace.getConfiguration('vbecm').get<boolean>('useFormModule');
  const isUseSheetModule = () => vscode.workspace.getConfiguration('vbecm').get<boolean>('useSheetModule');

  // // test   //vbatx.test
  // const commandTest = vscode.commands.registerCommand(
  //   'vbatx.test', 
  //   (uri:vscode.Uri) => {
  //     const path1 = 'C:\\projects\\toramame-hub\\xy\\xlsms\\src_macroTest.xlsm\\.base';
  //     const path2 = 'C:\\projects\\toramame-hub\\xy\\xlsms\\src_macroTest.xlsm\\.current';
  //     doDiff(path1, path2);
  //   });
  // context.subscriptions.push(commandTest);

  // export
  const commandExport = vscode.commands.registerCommand(
    'command.export', 
    async (uri:vscode.Uri) => {
      const exportModulesVbs = path.resolve(vbsPath, 'export.vbs');
      const xlsmFile = uri.fsPath;

      const fileDir = path.dirname(xlsmFile)
      const baseName = path.basename(xlsmFile)
      const srcDir = path.resolve(fileDir,'src_' + baseName)

      const isExist = await DirExists(srcDir)
      if (isExist){
        const ans = await vscode.window.showInformationMessage("Already exported. Do you want to export?", "Yes", "No")
        if (ans === 'No'){
          return
      }}

      const {err, status} = runVbs([exportModulesVbs, xlsmFile] );

      if (status !== 0 || err){
        vscode.window.showErrorMessage(err);
      } else {
        vscode.window.showInformationMessage("Success export");
      }
    });

  context.subscriptions.push(commandExport);

  // commit
  const commandImport = vscode.commands.registerCommand(
    'command.import', 
    (uri:vscode.Uri) => {
      console.log(uri.fsPath);
      const importModulesVbs = path.resolve(vbsPath, 'import.vbs');
      const xlsmFile = uri.fsPath;

      const isImportForm = isUseFormModule() ? 'True':'False';
      const isImportSheet = isUseSheetModule() ? 'True':'False';
      
      const {err, status} = runVbs([importModulesVbs, xlsmFile, isImportForm, isImportSheet] );

      if (status !== 0 || err){
        vscode.window.showErrorMessage(err);
      } else {
        vscode.window.showInformationMessage("Success import");
      }
    });
  context.subscriptions.push(commandImport);

  // compile
  const commandCompile = vscode.commands.registerCommand(
    'command.compile', 
    (uri:vscode.Uri) => {
      console.log(uri.fsPath);
      const compileVbs = path.resolve(vbsPath, 'compile.vbs');
      const xlsmFile = uri.fsPath;
  
      const {err, status} = runVbs([compileVbs, xlsmFile] );
  
      if (status !== 0 || err){
        vscode.window.showErrorMessage(err);
      } else {
        vscode.window.showInformationMessage("Success compile");
      }
    }
  );
  context.subscriptions.push(commandCompile);

  
  // run
  const commandRun = vscode.commands.registerTextEditorCommand (
    'editor.run', 
    async (textEditor: TextEditor, edit: vscode.TextEditorEdit, uri:vscode.Uri) => {
      // excel file path
      const dirVbeModule = path.basename(uri.fsPath).slice(".xlsm".length);
      const xlsmFile = await getExcelPathFromModule(uri);

      // sub function
      const activeLine = textEditor.document.lineAt(textEditor.selection.active.line).text;
      const vbaSub = activeLine.match(/Sub (.*)\(\s*\)/);
      const funcName = vbaSub !== null && vbaSub.length === 2 && vbaSub[1];
      const runModulesVbs = path.resolve(vbsPath, 'runVba.vbs');

      if (funcName){
        const {err, status} = runVbs([runModulesVbs, xlsmFile, funcName] );

        if (status !== 0 || err){
          vscode.window.showErrorMessage(err);
        } else {
          //vscode.window.showInformationMessage("file commit");
        }
      }
      else{
        console.log('No sub function is selected.')
      }
    });
  context.subscriptions.push(commandRun);

  // commit form editor
  const commandCommit = vscode.commands.registerCommand(
    'editor.commit', 
    async (uri:vscode.Uri) => {
      const importModulesVbs = path.resolve(vbsPath,'import.vbs');
      const xlsmFile = await getExcelPathFromModule(uri);
      
      const isImportForm = isUseFormModule() ? 'True':'False';
      const isImportSheet = isUseSheetModule() ? 'True':'False';
      
      const {err, status} = runVbs([importModulesVbs, xlsmFile, isImportForm, isImportSheet] );

      if (status !== 0){
        vscode.window.showErrorMessage(err);
      } else {
        vscode.window.showInformationMessage("file commit");
      }
    });
}

//
const getExcelPathFromModule = async (uri:vscode.Uri)  => {
  const dirParent = path.dirname(uri.fsPath);
  const dirForBook = path.dirname(dirParent);
  const bookFileName = path.basename(path.dirname(uri.fsPath)).slice("src_".length);
  const xlsPath = path.resolve(dirForBook, bookFileName);

  const isFile = await fileExists(xlsPath);

  return isFile ? xlsPath : '';
};

// this method is called when your extension is deactivated
export function deactivate() {}

import * as fs from 'fs' 
async function fileExists(filepath: string) {
  try {
    const res = (await fs.promises.lstat(filepath)).isFile();
    return (res)
  } catch (e) {
    return false
  }
}

async function DirExists(filepath: string) {
  try {
    const res = (await fs.promises.lstat(filepath)).isDirectory();
    return (res)
  } catch (e) {
    return false
  }
}

import * as iconv from 'iconv-lite';

// run vbs
const runVbs = (param:string[]) =>{
  console.log(param);
  const spawn = require( 'child_process').spawnSync,
  vbs = spawn( 'cscript.exe', ['//Nologo', ...param] );
  const err = s2u(vbs.stderr);
  const out = s2u(vbs.stdout);
  console.log( `stderr: ${err}` );
  console.log( `stdout: ${out}` );
  console.log( `status: ${vbs.status}` );

  return {err, out,  status: vbs.status};
};

// shift jis 2 utf8
const s2u = (sb: Buffer) =>{
  return iconv.decode(sb, "windows-31j");
};