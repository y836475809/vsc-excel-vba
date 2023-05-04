import * as vscode from 'vscode';

export class MyTreeItem extends vscode.TreeItem {
    vbaCommand: string;
    constructor(label: string, vbaCommand: string) {
        super(label);
        this.vbaCommand = vbaCommand;
    }
}

export class TreeDataProvider implements vscode.TreeDataProvider<vscode.TreeItem> {
    getTreeItem(element: vscode.TreeItem): vscode.TreeItem {
        return element;
    }

    getChildren(element?: vscode.TreeItem): vscode.TreeItem[] {
        return [
            new MyTreeItem("GotoVSCode", "gotoVSCode"),
            new MyTreeItem("GotoVBA", "gotoVBA"),
            new MyTreeItem("Import", "import"),
            new MyTreeItem("Export", "export"),
            new MyTreeItem("Compile", "compile"),  
        ];
    }
}