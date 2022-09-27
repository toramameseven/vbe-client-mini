// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import path = require('path');
import * as vscode from 'vscode';
import * as statusBar from './statusBar';
import * as VbeOutput from './vbeOutput';
import * as handler from './handlers';
import * as vbecmCommon from './vbecmCommon';

export const vbeReadOnlyDocumentProvider = new vbecmCommon.VbecmTextDocumentContentProvider();

// this method is called when your extension is activated
// your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {
  console.log('extension "vbe-client-mini" is now start!');

  // add status bar
  context.subscriptions.push(statusBar.statusBarVba);

  // register a content provider for the vbecm-scheme
  // for create readonly module
  const vbeScheme = 'vbecm';
  context.subscriptions.push(
    vscode.workspace.registerTextDocumentContentProvider(vbeScheme, vbeReadOnlyDocumentProvider)
  );

  //   --------------book
  context.subscriptions.push(
    vscode.commands.registerCommand('book.export', handler.handlerExportModulesFromBook)
  );

  context.subscriptions.push(
    vscode.commands.registerCommand('book.import', handler.handlerImportModulesToBook)
  );

  context.subscriptions.push(
    vscode.commands.registerCommand('book.compile', handler.handlerCompile)
  );

  context.subscriptions.push(
    vscode.commands.registerCommand('book.exportFrx', handler.handlerExportFrxModulesFromBook)
  );

  // push all module from folder
  context.subscriptions.push(
    vscode.commands.registerCommand(
      'srcFolder.commit-all',
      handler.handlerCommitAllModuleFromFolder
    )
  );

  // pull all module from folder
  context.subscriptions.push(
    vscode.commands.registerCommand(
      'srcFolder.pull-all',
      handler.handlerExportModulesFromFolder
    )
  );

  context.subscriptions.push(
    vscode.commands.registerCommand('srcFolder.checkModified', handler.handlerCheckModifiedOnFolder)
  );

  context.subscriptions.push(
    vscode.commands.registerCommand('srcFolder.compile', handler.handlerCompileFolder)
  );


  //   --------------editor
  context.subscriptions.push(
    vscode.commands.registerTextEditorCommand('editor.run', handler.handlerVbaRunFromEditor)
  );

  // check out on editor
  context.subscriptions.push(
    vscode.commands.registerCommand('editor.pullModule', handler.handlerPullModuleAsync)
  );

  // commit form editor
  context.subscriptions.push(
    vscode.commands.registerCommand('editor.commit', handler.handlerCommitModuleFromFile)
  );

  // goto vbe
  context.subscriptions.push(
    vscode.commands.registerTextEditorCommand('editor.gotoVbe', handler.handlerGotoVbe)
  );

  //  /////////////////////////////////  //////////////////////////////////////////////////
  context.subscriptions.push(
    vscode.commands.registerCommand(
      'vbeDiffView.refreshModification',
      handler.handlerUpdateModificationOnFolder
    )
  );

  context.subscriptions.push(
    vscode.commands.registerCommand(
      'vbeDiffView.collapseAll',
      handler.handlerCollapseAllVbeDiffView
    )
  );
  context.subscriptions.push(
    vscode.commands.registerCommand('vbeDiffView.pushSrc', handler.handlerCommitModuleFromVbeDiff)
  );

  context.subscriptions.push(
    vscode.commands.registerCommand('vbeDiffView.resolveVbe', (info) =>
      handler.handlerResolveVbeConflicting(info)
    )
  );

  context.subscriptions.push(
    vscode.commands.registerCommand('vbeDiffView.diffBaseTo', (resource) =>
      handler.handlerDiffBaseTo(resource)
    )
  );
  context.subscriptions.push(
    vscode.commands.registerCommand('vbeDiffView.diffVbeSrc', (resource) =>
      handler.handlerDiffSrcToVbe(resource)
    )
  );

  context.subscriptions.push(
    vscode.workspace.onDidSaveTextDocument((event) => {
      //vbeOutput.appendLine(`onDidSaveTextDocument: ${event.fileName}`);
      handler.handlerCheckModifiedOnSave(event.uri);
    })
  );

  displayMenus(true);

  // update status bar item once at start
  statusBar.updateStatusBarItem(false);

  //log
  VbeOutput.vbeOutput.show(false);
  VbeOutput.showInfo('vbecm extension start', false);
}

//End ------------------------------------------------------------------------

// this method is called when your extension is deactivated
export function deactivate() {}

/**
 * set display vbe menu on or off
 * @param isOn
 */
const displayMenus = (isOn: boolean) => {
  vscode.commands.executeCommand('setContext', 'vbecm.showVbsCommand', isOn);
  statusBar.updateStatusBarItem(!isOn);
};
