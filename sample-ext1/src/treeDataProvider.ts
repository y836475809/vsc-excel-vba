import * as vscode from 'vscode';

export class TreeDataProvider implements vscode.TreeDataProvider<vscode.TreeItem> {
    getTreeItem(element: vscode.TreeItem): vscode.TreeItem {
        return element;
    }

    getChildren(element?: vscode.TreeItem): vscode.TreeItem[] {
        return [
            new vscode.TreeItem('test1'), 
            new vscode.TreeItem('test2'), 
            new vscode.TreeItem('test3')
        ];
    }
}