import * as fs from "fs";
import * as path from "path";
import * as vscode from "vscode";

export class Project {
    async setupConfig() {
        const config = vscode.workspace.getConfiguration();
        const editorConfig: any = await config.get("editor.quickSuggestions");
        if(editorConfig["other"] !== "off"){
            editorConfig["other"] = "on";
            await config.update(
                "editor.quickSuggestions",
                editorConfig,
                vscode.ConfigurationTarget.Workspace
            );
        }
        await config.update("files.autoGuessEncoding", true);
    }

    async copy(sourcePath:string, destPath:string){
        const wsedit = new vscode.WorkspaceEdit();
        const data = await vscode.workspace.fs.readFile(
            vscode.Uri.file(sourcePath)
        );
        const filePath = vscode.Uri.file(destPath);
        wsedit.createFile(filePath, { ignoreIfExists: true });
        await vscode.workspace.fs.writeFile(filePath, data);
        let isDone = await vscode.workspace.applyEdit(wsedit);
    }
}