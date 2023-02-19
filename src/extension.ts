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
    vscode.commands.registerCommand('book.export', handler.handlerBookExportModules)
  );

  context.subscriptions.push(
    vscode.commands.registerCommand('book.import', handler.handlerBookImportModules)
  );

  context.subscriptions.push(
    vscode.commands.registerCommand('book.compile', handler.handlerBookCompile)
  );

  context.subscriptions.push(
    vscode.commands.registerCommand('book.exportFrx', handler.handlerBookExportFrxModules)
  );

  // push all module from folder
  context.subscriptions.push(
    vscode.commands.registerCommand(
      'srcFolder.commit-all',
      handler.handlerSrcFolderCommitAllModules
    )
  );

  // pull all module from folder
  context.subscriptions.push(
    vscode.commands.registerCommand('srcFolder.pull-all', handler.handlerSrcFolderExportModulesTo)
  );

  context.subscriptions.push(
    vscode.commands.registerCommand(
      'srcFolder.checkModified',
      handler.handlerSrcFolderRefreshModification
    )
  );

  context.subscriptions.push(
    vscode.commands.registerCommand('srcFolder.compile', handler.handlerSrcFolderCompile)
  );

  //   --------------editor
  context.subscriptions.push(
    vscode.commands.registerTextEditorCommand('editor.run', handler.handlerEditorVbaRun)
  );

  // check out on editor
  context.subscriptions.push(
    vscode.commands.registerCommand('editor.pullModule', handler.handlerEditorPullModule)
  );

  // commit form editor
  context.subscriptions.push(
    vscode.commands.registerCommand('editor.commit', handler.handlerEditorCommitModule)
  );

  // commit form editor only frm. a frx in the vbe is used.
  // context.subscriptions.push(
  //   vscode.commands.registerCommand(
  //     'editor.commitOnlyFrm',
  //     handler.handlerEditorCommitModuleOnlyFrm
  //   )
  // );

  context.subscriptions.push(
    vscode.commands.registerCommand('editor.pullOnlyFrx', handler.handlerEditorPullModuleFrx)
  );

  // goto vbe
  context.subscriptions.push(
    vscode.commands.registerTextEditorCommand('editor.gotoVbe', handler.handlerEditorGotoVbe)
  );

  //  /////////////////////////////////  //////////////////////////////////////////////////
  context.subscriptions.push(
    vscode.commands.registerCommand(
      'vbeDiffView.refreshModification',
      handler.handlerDiffViewRefreshModification
    )
  );

  context.subscriptions.push(
    vscode.commands.registerCommand('vbeDiffView.collapseAll', handler.handlerDiffViewCollapseAll)
  );
  context.subscriptions.push(
    vscode.commands.registerCommand('vbeDiffView.pushSrc', handler.handlerDiffViewCommitModule)
  );

  context.subscriptions.push(
    vscode.commands.registerCommand('vbeDiffView.pullSrc', handler.handlerDiffViewPullModule)
  );

  // delete this command, for very confused
  // context.subscriptions.push(
  //   vscode.commands.registerCommand('vbeDiffView.resolveVbe', (info) =>
  //     handler.handlerResolveVbeConflicting(info)
  //   )
  // );

  // define at tree
  context.subscriptions.push(
    vscode.commands.registerCommand('vbeDiffView.diffBaseTo', (resource) =>
      handler.handlerDiffViewDiffBaseTo(resource)
    )
  );

  // add tree for conflict
  // context.subscriptions.push(
  //   vscode.commands.registerCommand('vbeDiffView.diffVbeSrc', (resource) =>
  //     handler.handlerDiffViewDiffSrcToVbe(resource)
  //   )
  // );

  context.subscriptions.push(
    vscode.workspace.onDidSaveTextDocument((event) => {
      const extensions = ['.bas', '.frm', '.cls', '.frx', '.vbs'];
      const extension = event.fileName.slice(-4).toLowerCase();
      if (!extensions.includes(extension)) {
        return;
      }
      handler.handlerOnSaveRefreshModification(event.uri);
    })
  );

  context.subscriptions.push(
    vscode.workspace.onDidOpenTextDocument((event) => {
      const extensions = ['.bas', '.frm', '.cls', '.frx', '.vbs'];
      const extension = event.fileName.slice(-4).toLowerCase();
      if (!extensions.includes(extension)) {
        return;
      }
      handler.handlerOnOpenRefreshModification(event);
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
