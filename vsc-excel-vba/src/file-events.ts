import * as vscode from "vscode";
import * as path from "path";
import { VBALSRequest } from "./vba-ls-request";

function debounce(fn: any, interval: number){
    let timerId: any;
    return (e: any) => {
		if(timerId){
			clearTimeout(timerId);
		}
        timerId = setTimeout(() => {
            fn(e);
        }, interval);
    };
};

export class FileEvents {
    request: VBALSRequest;
    disposables: vscode.Disposable[];

    constructor(request: VBALSRequest){
        this.request = request;
        this.disposables = [];
    }

    public dispose() {
        this.disposables.forEach(x=>x.dispose());
        this.disposables = [];
    }

    public registerFileEvent(): void {
        this.dispose();

        this.disposables.push(vscode.workspace.onDidCreateFiles(async (e: vscode.FileCreateEvent) => {
            const uris = e.files.map(x => x);
            await this.request.addDocuments(uris);
        }));

        this.disposables.push(vscode.workspace.onDidDeleteFiles(async (e: vscode.FileDeleteEvent) => {
            const uris = e.files.map(x => x);
            await this.request.deleteDocuments(uris);
        }));

        this.disposables.push(vscode.workspace.onDidRenameFiles(async (e: vscode.FileRenameEvent) => {
            const files = e.files.map(file => {
                return {
                    oldUri: file.oldUri,
                    newUri: file.newUri
                };
            });
            await this.request.renameDocument(files);
        }));

        const delayTimeMs = 1000;
        this.disposables.push(vscode.workspace.onDidChangeTextDocument(
            debounce(async (e: vscode.TextDocumentChangeEvent) => {
                if(e.document.isUntitled){
                    return;
                }
                // if(!e.document.isDirty && !e.reason){
                //     // 変更なしでsave
                //     return;
                // }
                // // isDirty=false, e.reason!=undefined
                // // ->undo or redoで変更をもどした場合なので更新通知必要

                await this.request.changeDocument(e.document);
        }, delayTimeMs)));

        this.disposables.push(vscode.window.onDidChangeActiveTextEditor(
            debounce(async (e: vscode.TextEditor) => { 
                const fname = e.document.fileName;
                if(fname.endsWith(".bas") || fname.endsWith(".cls")){
                    await this.request.diagnostic(e.document);
                }
        }, delayTimeMs)));
    }
} 