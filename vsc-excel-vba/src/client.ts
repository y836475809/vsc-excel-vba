import * as path from "path";
import * as vscode from 'vscode';
import {
	LanguageClient,
	LanguageClientOptions,
	ServerOptions,
	TransportKind,
} from 'vscode-languageclient/node';


export function createClient(context: vscode.ExtensionContext, serverPath: string, srcDir: string): LanguageClient {
    const srcDirName = path.basename(srcDir);
    const serverOptions: ServerOptions = {
        run: {
            command: serverPath,
            args: [`--src_dir_name=${srcDirName}`],
            transport: TransportKind.stdio,
        },
        debug: {
            command: serverPath,
            args: [`--src_dir_name=${srcDirName}`],
            transport: TransportKind.stdio,
        }	
     };

    const clientOptions: LanguageClientOptions = {
        documentSelector: [{ scheme: "file", language: "vba" }],
        synchronize: {
            fileEvents: vscode.workspace.createFileSystemWatcher(`${srcDir}/*.{bas,cls}`)
        },
    };
    const client = new LanguageClient(
        "VBALanguageServer",
        "VBA Language Server",
        serverOptions,
        clientOptions
    );
    vscode.window.withProgress(
        {
            location: vscode.ProgressLocation.Notification, 
            title: "Initializing VBA Language Server"
        }, async progress => {
            return new Promise<void>((resolve)=>{
                const listener = client.onNotification("custom/initialized", e => {
                    resolve();
                });
                context.subscriptions.push(listener);
            });
        });
    return client;
};