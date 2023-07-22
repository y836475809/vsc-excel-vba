import * as vscode from 'vscode';

export class SheetTreeItem extends vscode.TreeItem {
    sheetName: string;

    constructor(sheetName: string) {
        super(sheetName);
        this.sheetName = sheetName;
        this.iconPath = new vscode.ThemeIcon("file");
    }
}

export class SheetTreeDataProvider implements vscode.TreeDataProvider<SheetTreeItem>, vscode.Disposable {
    private sheets: SheetTreeItem[];
    private treeview: vscode.TreeView<vscode.TreeItem>;

    private _onDidChangeTreeData: vscode.EventEmitter<void> = new vscode.EventEmitter<void>();
    readonly onDidChangeTreeData: vscode.Event<void> = this._onDidChangeTreeData.event;

    constructor(viewId: string){
        this.sheets = [];
        this.treeview = vscode.window.createTreeView(viewId, {
            treeDataProvider: this,
            showCollapseAll: true,
            canSelectMany: false
        });
    }

    dispose() {
        this.sheets = [];
        this._onDidChangeTreeData.fire();
        this.treeview.dispose();
    }

    refresh(sheets: string[]): void {
        this.sheets = sheets.map(x => {
            return new SheetTreeItem(x);
        });
       
        this._onDidChangeTreeData.fire();
    }

    getTreeItem(element: vscode.TreeItem): vscode.TreeItem {
        return element;
    }

    getChildren(element?: vscode.TreeItem): SheetTreeItem[] {
        return this.sheets;
    }
}