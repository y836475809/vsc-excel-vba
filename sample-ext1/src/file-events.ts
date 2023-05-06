import {
	LanguageClient,
	State
} from 'vscode-languageclient/node';
import * as vscode from "vscode";
import * as path from "path";

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
    client: LanguageClient;
    srcDir: string;

    constructor(client: LanguageClient, srcDir: string){
        this.client = client;
        this.srcDir = srcDir;
    }

    private isInSrcDir(uri: vscode.Uri): boolean{
        return path.dirname(uri.fsPath) === this.srcDir;
    }

    private async renameFiles(files: any[]){
        const method: Hoge.RequestMethod = "renameFiles";
        let renameParams: Hoge.RequestRenameParam[] = [];
        for(const file of files){
            const oldUri = file.oldUri.toString();
            const newUri = file.newUri.toString();
            renameParams.push({
                olduri: oldUri,
                newuri: newUri
            });
        }
        await this.client.sendRequest(method, {renameParams});
    }

    private async deleteFiles(files: any[]){
        const method: Hoge.RequestMethod = "deleteFiles";
        const uris = files.map(uri => {
            return uri.toString();
        });
        await this.client.sendRequest(method, {uris});
    }

    public registerFileEvent(context: vscode.ExtensionContext): vscode.Disposable[] {
        const feDisps: vscode.Disposable[] = [];

        feDisps.push(vscode.workspace.onDidCreateFiles(async (e: vscode.FileCreateEvent) => {
            if(!this.client || this.client.state !== State.Running){
                return;
            }
            const method: Hoge.RequestMethod = "createFiles";
            const uris = e.files.filter(file => this.isInSrcDir(file)).map(uri => uri.toString());
            if(!uris.length){
                return;
            }
            await this.client.sendRequest(method, {uris});
        }, null, context.subscriptions));

        feDisps.push(vscode.workspace.onDidDeleteFiles(async (e: vscode.FileDeleteEvent) => {
            if(!this.client || this.client.state !== State.Running){
                return;
            }
            const files = e.files.filter(file => this.isInSrcDir(file)).map(file => {
                return file;
            });
            if(!files.length){
                return;
            }
            await this.deleteFiles(files);
        }, null, context.subscriptions));

        feDisps.push(vscode.workspace.onDidRenameFiles(async (e: vscode.FileRenameEvent) => {
            if(!this.client || this.client.state !== State.Running){
                return;
            }
            const files = e.files.filter(file => this.isInSrcDir(file.newUri)).map(file => {
                return {
                    oldUri: file.oldUri,
                    newUri: file.newUri
                };
            });
            if(!files.length){
                return;
            }
            await this.renameFiles(files);
        }, null, context.subscriptions));

        feDisps.push(vscode.workspace.onDidChangeTextDocument(
            debounce(async (e: vscode.TextDocumentChangeEvent) => {
                if(!this.client || this.client.state !== State.Running){
                    return;
                }
                if(e.document.isUntitled){
                    return;
                }
                if(!e.document.isDirty && !e.reason){
                    // 変更なしでsave
                    return;
                }
                // isDirty=false, e.reason!=undefined
                // ->undo or redoで変更をもどした場合なので更新通知必要

                if(!this.isInSrcDir(e.document.uri)){
                    return;
                }
                const method: Hoge.RequestMethod = "changeText";
                const uri = e.document.uri.toString();
                await this.client.sendRequest(method, {uri});
        }, 500), null, context.subscriptions));

        feDisps.push(vscode.window.onDidChangeActiveTextEditor(
            debounce(async (e: vscode.TextEditor) => {
                if(!this.client || this.client.state !== State.Running){
                    return;
                }
                const fname = e.document.fileName;
                if(fname.endsWith(".bas") || fname.endsWith(".cls")){
                    const method: Hoge.RequestMethod = "diagnostics";
                    await this.client.sendRequest(method, {uri:e.document.uri.toString()});
                }
        }, 1000), null, context.subscriptions));

        return feDisps;
    }
} 