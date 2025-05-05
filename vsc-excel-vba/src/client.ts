import * as path from "path";
import * as vscode from 'vscode';
import {
	LanguageClient,
	LanguageClientOptions,
	ServerOptions,
	TransportKind,
} from 'vscode-languageclient/node';


export function vbaClient(context: vscode.ExtensionContext, srcDir: string): LanguageClient {
    const srcDirName = path.basename(srcDir);
    const config = vscode.workspace.getConfiguration();
    const lspFilename = config.get("vsc-excel-vba.LSFilename") as string;
    const serverOptions: ServerOptions = {
        run: {
            command: path.join(context.asAbsolutePath("bin"), "Release", lspFilename),
            args: [`--src_dir_name=${srcDirName}`],
            transport: TransportKind.stdio,
        },
        debug: {
            command: path.join(context.asAbsolutePath("bin"), "Debug", lspFilename),
            args: [`--src_dir_name=${srcDirName}`],
            transport: TransportKind.stdio,
        }	
     };

    const clientOptions: LanguageClientOptions = {
        documentSelector: [{ scheme: 'file', language: 'vb' }],
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