// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import path = require('path');
import * as vscode from 'vscode';
import * as statusBar from './statusBar';
import * as vbs from './vbsModule';
import { vbeOutput } from './vbeOutput';
import * as handler from './handlers';



// this method is called when your extension is activated
// your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {
  
  console.log('extension "vbe-client-mini" is now active!');

  //statusBar.statusBarVba = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Right, 100);
	// statusBarVba.command = myCommandId;
	context.subscriptions.push(statusBar.statusBarVba);

  // const isUseFormModule = () => vscode.workspace.getConfiguration('vbecm').get<boolean>('useFormModule');
  // const isUseSheetModule = () => vscode.workspace.getConfiguration('vbecm').get<boolean>('useSheetModule');
  
  //   --------------explorer
  // export
  const commandExport = vscode.commands.registerCommand(
    'command.export', 
    handler.handlerExportModules
  );
  context.subscriptions.push(commandExport);

  // import
  const commandImport = vscode.commands.registerCommand(
    'command.import', 
    handler.handlerImportModules
  );
  context.subscriptions.push(commandImport);

  // compile
  const commandCompile = vscode.commands.registerCommand(
    'command.compile', 
    handler.handlerCompile
  );
  context.subscriptions.push(commandCompile);

  // export frx
  const commandExportFrx = vscode.commands.registerCommand(
    'command.exportfrx', 
    handler.handlerUpdateFrxModules
  );
  context.subscriptions.push(commandExportFrx);

  // commit all module from folder
  const commandCommitAll = vscode.commands.registerCommand(
    'command.commit-all', 
    handler.handlerCommitAllModule
  );
  context.subscriptions.push(commandCommitAll);

  //   --------------editor
  // run
  const commandRun = vscode.commands.registerTextEditorCommand (
    'editor.run', 
    handler.handlerVbaRun
  );
  context.subscriptions.push(commandRun);

  // check out on editor
  const commandCheckOut = vscode.commands.registerCommand(
    'editor.checkout', 
    handler.handlerUpdateAsync
  );
  context.subscriptions.push(commandCheckOut);

  // commit form editor
  const commandCommit = vscode.commands.registerCommand(
    'editor.commit', 
    handler.handlerCommitModule
  );
  context.subscriptions.push(commandCommit);








  
  displayMenus(true);
  // update status bar item once at start
	statusBar.updateStatusBarItem(false);
  
  //log
  vbeOutput.show(false);
  vbeOutput.appendLine('initialize output');

}

//End ------------------------------------------------------------------------

// this method is called when your extension is deactivated
export function deactivate() {}

/**
 * set display vbe menu on or off
 * @param isOn
 */
const displayMenus = (isOn : boolean) =>{
  vscode.commands.executeCommand('setContext', 'vbecm.showVbsCommand', isOn);
  statusBar.updateStatusBarItem(!isOn);
};


