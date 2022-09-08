// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import path = require('path');
import * as vscode from 'vscode';
import * as statusBar from './statusBar';
import * as vbs from './vbsModule';
import { vbeOutput } from './vbeOutput';
import * as handler from './handlers';
import * as fse from 'fs-extra';
import * as encoding from 'encoding-japanese';
import { FileDiff, FileDiffProvider } from './diffFiles';


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

  const commandModified = vscode.commands.registerCommand(
    'command.modified', 
    handler.handlerCheckModified
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
  
  const fileDiffProvider = new FileDiffProvider(['aaaa','bbbb'], 'src');
	vscode.window.registerTreeDataProvider('vbeDiff', fileDiffProvider);

  const commandDiffRefresh = vscode.commands.registerCommand('vbeDiff.refreshEntry', () => fileDiffProvider.refresh());
  context.subscriptions.push(commandDiffRefresh);

  
  //

  // //when open file check japanese encode
	// vscode.workspace.onDidOpenTextDocument(function(e){
	// 	const fname = e.fileName.replace(/\\/g,'/');
	// 	if (fname.indexOf('/.vscode') > -1){
	// 		return;
	// 	}

	// 	//workspaceデフォルトのエンコードを取得
	// 	const  defaultEncode = vscode.workspace.getConfiguration().get('files.encoding');// utf8 shiftjis eucjp

  //   //vscode.window.activeTextEditor?.document.


	// 	//元ファイルを開いてエンコードを調べる
	// 	fse.readFile(fname, function (err, textx) {
  //     const text = e.getText();
	// 		if (text.length === 0){
	// 			return;
	// 		}
	// 		const sourceEncode = encoding.detect(text);// UTF8 SJIS EUCJP ASCII
	// 		if (!sourceEncode){
	// 			return;
	// 		}

	// 		const sourceEncodeLowerCase = sourceEncode.toLowerCase() === 'sjis' ? 'shiftjis': sourceEncode.toLowerCase();

	// 		let mes = STRING_EMPTY;
	// 		if (sourceEncodeLowerCase !== defaultEncode && sourceEncodeLowerCase !== 'ascii'){
	// 			mes = 'encoding not match!! reopen with [' + sourceEncode + '] ' + fname;
	// 			vscode.window.showWarningMessage(mes);
	// 		}else{
	// 			mes = 'encoding match with workspace default(' + defaultEncode + '). [' + sourceEncode + '] ' + e.fileName;
	// 			vscode.window.setStatusBarMessage(mes,5000);
	// 		}
	// 	});
	// });	

  
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




