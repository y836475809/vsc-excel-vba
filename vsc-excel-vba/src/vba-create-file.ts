import * as vscode from 'vscode';
import * as path from "path";
import * as fs from "fs";
import { VBAObjectNameValidation } from "./vba-object-name-validation";

export class VBACreateFile {
    private objectType: VEV.VBAObjectType;

    constructor(){
        this.objectType = "class";
    }

    async create(distUri: vscode.Uri, objectType: VEV.VBAObjectType) {
        this.objectType = objectType;
        const validate = new VBAObjectNameValidation();
        const result = await vscode.window.showInputBox({
            prompt: `Input ${this.objectType} name`,
            value: `${this.objectType}1`,
            validateInput: (text) => {
                const filename = this.getFileName(text);
                const filePath = path.join(distUri.fsPath, filename);
                if(fs.existsSync(filePath)){
                    return `${filename} exists`;
                }
                
                if(!validate.prefix(text)){
                    return "Name contains illegal prefix";
                }
                if(!validate.symbol(text)){
                    return "Name contains illegal symbols";
                }
                if(!validate.len(text)){
                    return "Name is too long";
                }
                return null;
            }
        });
        if (result) {
            const filename = this.getFileName(result);
            const newUri = vscode.Uri.joinPath(distUri, filename);
            const untitledUri = vscode.Uri.parse(`untitled:${newUri.fsPath}`);
            const doc = await vscode.workspace.openTextDocument(untitledUri);
            const contentLines = this.getContentLines(result);
            const lineNum = contentLines.length;
            const editor = await vscode.window.showTextDocument(doc, vscode.ViewColumn.Active, true);
            await editor.edit(edit => {
                edit.insert(new vscode.Position(0, 0), contentLines.join("\r\n"));
            });
            await doc.save();
            await vscode.window.showTextDocument(newUri, { 
                preview: false,
                selection: new vscode.Range(
                    new vscode.Position(lineNum, 0),
                    new vscode.Position(lineNum, 0))
            });
        }
    }

    private getContentLines(name: string){
        if(this.objectType === "module"){
            return [
                `Attribute VB_Name = "${name}"`,
                `Option Explicit`,
                ``
            ];
        }
        if(this.objectType === "class"){
            return [
                `VERSION 1.0 CLASS`,
                `BEGIN`,
                `  MultiUse = -1`,
                `END`,
                `Attribute VB_Name = "${name}"`,
                `Attribute VB_GlobalNameSpace = False`,
                `Attribute VB_Creatable = False`,
                `Attribute VB_PredeclaredId = False`,
                `Attribute VB_Exposed = False`,
                `Option Explicit`,
                ``
            ];
        }

        throw new Error(`${this.objectType} is not supported`);
    }

    private getFileName(name: string): string {
        const ext = this.getExt();
        return `${name}.${ext}`;
    }

    private getExt(){
        if(this.objectType === "class"){
            return "cls";
        }
        if(this.objectType === "module"){
            return "bas";
        }
    }
}
