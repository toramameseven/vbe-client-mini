// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import path = require('path');
import * as vscode from 'vscode';
import * as statusBar from './statusBar';
import * as vbs from './vbsModule';

// this method is called when your extension is activated
// your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {
  
  console.log('Congratulations, your extension "vbe-client-mini" is now active!');

  //statusBar.statusBarVba = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Right, 100);
	// statusBarVba.command = myCommandId;
	context.subscriptions.push(statusBar.statusBarVba);

  // const isUseFormModule = () => vscode.workspace.getConfiguration('vbecm').get<boolean>('useFormModule');
  // const isUseSheetModule = () => vscode.workspace.getConfiguration('vbecm').get<boolean>('useSheetModule');
  
  // export
  const commandExport = vscode.commands.registerCommand(
    'command.export', 
    vbs.exportModuleAsync
  );
  context.subscriptions.push(commandExport);

  // import
  const commandImport = vscode.commands.registerCommand(
    'command.import', 
    vbs.importModuleAsync
  );
  context.subscriptions.push(commandImport);

  // compile
  const commandCompile = vscode.commands.registerCommand(
    'command.compile', 
    vbs.compile
  );
  context.subscriptions.push(commandCompile);

  // run
  const commandRun = vscode.commands.registerTextEditorCommand (
    'editor.run', 
    vbs.runAsync
  );
  context.subscriptions.push(commandRun);

  // check out on editor
  const commandCheckOut = vscode.commands.registerCommand(
    'editor.checkout', 
    vbs.checkoutAsync
  );
  context.subscriptions.push(commandCheckOut);

  // commit form editor
  const commandCommit = vscode.commands.registerCommand(
    'editor.commit', 
    vbs.commitAsync
  );
  context.subscriptions.push(commandCommit);
  
  displayMenu(true);
  // update status bar item once at start
	statusBar.updateStatusBarItem(false);
}

//End ------------------------------------------------------------------------

// this method is called when your extension is deactivated
export function deactivate() {}

/**
 * set display vbe menu on or off
 * @param isOn
 */
const displayMenu = (isOn : boolean) =>{
  vscode.commands.executeCommand('setContext', 'vbecm.showVbsCommand', isOn);
  statusBar.updateStatusBarItem(!isOn);
};


