import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import { commands } from 'vscode';
import { getVbecmDirs, STRING_EMPTY , compareModules, DiffTitle, DiffState} from './vbsModule';
import { myLog, displayMenus } from './common';
import * as common from './common';

export type DiffFileInfo  = {
  isHeader?: boolean;
  moduleName?: string;
  diffState?: 'modified' | 'add' | 'removed';
  isBase?:boolean;
  baseFilePath?: string;
  isCompare?:boolean;
  compareFilePath?: string;
  titleBaseTo?: DiffTitle;
};

const isOut = true;

export class FileDiffProvider implements vscode.TreeDataProvider<DiffFileInfo> {
  refresh(): any {
    myLog('run refresh','refresh', isOut);
    this._onDidChangeTreeData.fire();
  }

	private _onDidChangeTreeData: vscode.EventEmitter<DiffFileInfo | undefined | void> = new vscode.EventEmitter<DiffFileInfo | undefined | void>();
	readonly onDidChangeTreeData: vscode.Event<DiffFileInfo | undefined | void> = this._onDidChangeTreeData.event;

	constructor(
    private bookPath: string,
    private filesBaseSrc: DiffFileInfo[] = [], 
    private filesBaseVbe: DiffFileInfo[] = [],
    ) {}

  setFiles(bookPath: string, filesBaseSrc: DiffFileInfo[], filesBaseVbe: DiffFileInfo[]){
    this.bookPath = bookPath;
    this.filesBaseSrc = filesBaseSrc;
    this.filesBaseVbe = filesBaseVbe;
  }

	getTreeItem(element: DiffFileInfo): vscode.TreeItem {
    if (element.isHeader ){
      const treeItem =  new FileDiff(element.moduleName ?? STRING_EMPTY, vscode.TreeItemCollapsibleState.Collapsed);
      return treeItem;
    } else {
      // files
      const treeItem =  new FileDiff(element.moduleName ?? STRING_EMPTY, vscode.TreeItemCollapsibleState.None);
      treeItem.command = { command: 'vbeDiffView.diffBaseTo', title: 'Open File', arguments: [element], };
      if (element.titleBaseTo === 'base-vbe: ') {
        treeItem.contextValue = 'FileDiffTreeItemBaseVbe';
      }
      treeItem.description = element.diffState;
      return treeItem;
    }
	}

  async getChildren(element?: DiffFileInfo): Promise<DiffFileInfo[]> {
    if(element){
      const r = element.moduleName === 'src(base)' ? this.filesBaseSrc : this.filesBaseVbe;
      return r;
    } else {
      // root
      const n : DiffFileInfo = this.filesBaseSrc.length > 0 ? {moduleName: 'src(base)', isHeader: true,} : {};
      const n1 : DiffFileInfo = this.filesBaseVbe.length > 0 ? {moduleName: 'vbe(base)', isHeader: true} : {};

      const r = [n, n1].filter(_ => _.moduleName);
      return r;
    }
	}

  async updateModification(bookPath?: string){
    displayMenus(false);

    let bookPathToUpdate = bookPath;
    if (!bookPathToUpdate) {
      const activeTextEditor = vscode.window.activeTextEditor?.document.uri;
      bookPathToUpdate = await this.getExcelPathFromModule(activeTextEditor?.path);
    }

    if (bookPathToUpdate && await common.fileExists(bookPathToUpdate)){
      // update path
      this.bookPath = bookPathToUpdate;
      vbeTreeView.title = 'diff: ' + path.basename(bookPathToUpdate);
    }

    try {
      const {diffBaseSrc, diffBaseVbe} = await compareModules(this.bookPath);
      fileDiffProvider.setFiles(this.bookPath, diffBaseSrc, diffBaseVbe);
      fileDiffProvider.refresh();
    } catch (error) {
      throw error;
    } finally {
      displayMenus();
    }
  }

  async getExcelPathFromModule(pathModule: string | undefined) {
    if (pathModule === undefined){return pathModule;}
    const dirParent = path.dirname(pathModule);
    const r = await this.getExcelPathSrcFolder(vscode.Uri.file(dirParent));
    return r;
  }
  async getExcelPathSrcFolder(uriSrcFolder: vscode.Uri) {
    const dirForBook = path.dirname(uriSrcFolder.fsPath);
    const bookFileName = path.basename(uriSrcFolder.fsPath).slice('src_'.length);
    const xlsPath = path.resolve(dirForBook, bookFileName);
    const isFile = await common.fileExists(xlsPath);
    return isFile ? xlsPath : undefined;
  }
}

export const fileDiffProvider = new FileDiffProvider(STRING_EMPTY, [], []);
export const vbeTreeView = vscode.window.createTreeView('vbeDiffView', {treeDataProvider: fileDiffProvider});

export class FileDiff extends vscode.TreeItem {
	constructor(
		public readonly label: string,
		public readonly collapsibleState: vscode.TreeItemCollapsibleState,
		public command?: vscode.Command,
    public description?: 'modified' | 'add' | 'removed'| undefined
	) {
		super(label, collapsibleState);
		// this.tooltip = 'tooltisp';
		this.description = description;
	}
	contextValue = 'FileDiffTreeItemBaseSrc';
}
