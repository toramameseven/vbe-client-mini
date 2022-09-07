import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';

export class FileDiffProvider implements vscode.TreeDataProvider<FileDiff> {

	private _onDidChangeTreeData: vscode.EventEmitter<FileDiff | undefined | void> = new vscode.EventEmitter<FileDiff | undefined | void>();
	readonly onDidChangeTreeData: vscode.Event<FileDiff | undefined | void> = this._onDidChangeTreeData.event;

	constructor(private files: string[] | undefined) {
	}

	refresh(): void {
		this._onDidChangeTreeData.fire();
	}

	getTreeItem(element: FileDiff): vscode.TreeItem {
		return element;
	}

	getChildren(element?: FileDiff): Thenable<FileDiff[]> {
    if (!this.files) {
      Promise.resolve([]);
    }
    return Promise.resolve(this.files!.map(_ => new FileDiff(_, vscode.TreeItemCollapsibleState.None )));
	}
}

export class FileDiff extends vscode.TreeItem {

	constructor(
		public readonly label: string,
		public readonly collapsibleState: vscode.TreeItemCollapsibleState,
		public readonly command?: vscode.Command
	) {
		super(label, collapsibleState);

		// this.tooltip = `${this.label}-${this.version}`;
		// this.description = this.version;
	}

	// iconPath = {
	// 	light: path.join(__filename, '..', '..', 'resources', 'light', 'dependency.svg'),
	// 	dark: path.join(__filename, '..', '..', 'resources', 'dark', 'dependency.svg')
	// };

	contextValue = 'fileDiffs';
}
