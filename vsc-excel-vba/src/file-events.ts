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
    srcDir: string;
    disposables: vscode.Disposable[];

    constructor(request: VBALSRequest){
        this.request = request;
        this.srcDir = "";
        this.disposables = [];
    }

    public dispose() {
        this.disposables.forEach(x=>x.dispose());
        this.disposables = [];
    }

    private isInSrcDir(uri: vscode.Uri): boolean{
        return path.dirname(uri.fsPath).startsWith(this.srcDir);
    }

    private async renameFiles(files: any[]){
        const method: Hoge.RequestId = "RenameDocument";
        let renameParams: Hoge.RequestRenameParam[] = [];
        for(const file of files){
            const oldUri = file.oldUri;
            const newUri = file.newUri;
            renameParams.push({
                olduri: oldUri,
                newuri: newUri
            });
        }
        await this.request.renameDocument(renameParams);
    }

    private async deleteFiles(files: any[]){
        const method: Hoge.RequestId = "DeleteDocuments";
        const uris = files.map(uri => {
            return uri.toString();
        }).filter(x => x.scheme === "file");
        if(!uris.length){
            return;
        } 
        await this.request.deleteDocuments(uris);
    }

    public registerFileEvent(srcDir: string): void {
        this.dispose();

        this.srcDir = srcDir;

        this.disposables.push(vscode.workspace.onDidCreateFiles(async (e: vscode.FileCreateEvent) => {
            const uris = e.files.filter(file => this.isInSrcDir(file))
                .filter(x => x.scheme === "file");
            if(!uris.length){
                return;
            }
            await this.request.addDocuments(uris);
        }));

        this.disposables.push(vscode.workspace.onDidDeleteFiles(async (e: vscode.FileDeleteEvent) => {
            const files = e.files.filter(file => this.isInSrcDir(file)).map(file => {
                return file;
            }).filter(x => x.scheme === "file");
            if(!files.length){
                return;
            }
            await this.deleteFiles(files);
        }));

        this.disposables.push(vscode.workspace.onDidRenameFiles(async (e: vscode.FileRenameEvent) => {
            const files = e.files.filter(file => this.isInSrcDir(file.newUri))
            .filter(x => x.newUri.scheme === "file" && x.oldUri.scheme === "file")
            .map(file => {
                return {
                    oldUri: file.oldUri,
                    newUri: file.newUri
                };
            });
            if(!files.length){
                return;
            }
            await this.renameFiles(files);
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
                const uri = e.document.uri;
                if(!this.isInSrcDir(uri)){
                    return;
                }
                if(uri.scheme !== "file"){
                    return;
                }      
                await this.request.changeDocument(e.document);
        }, delayTimeMs)));

        this.disposables.push(vscode.window.onDidChangeActiveTextEditor(
            debounce(async (e: vscode.TextEditor) => {
                if(e.document.uri.scheme !== "file"){
                    return;
                }  
                const fname = e.document.fileName;
                if(fname.endsWith(".bas") || fname.endsWith(".cls")){
                    await this.request.diagnostic(e.document);
                }
        }, delayTimeMs)));
    }
} 