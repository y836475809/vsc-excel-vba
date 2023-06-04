import * as path from 'path';
import * as vscode from 'vscode';
import {
	LanguageClient,
	LanguageClientOptions,
	ServerOptions,
	State,
	TransportKind
} from 'vscode-languageclient/node';
import { FileEvents } from "./file-events";

export class LSPClient {
	client!: LanguageClient;
	wsFileEventDisps: vscode.Disposable[];

	constructor(){
		this.wsFileEventDisps = [];
	}

	registerFileEvents(srcDir: string){
		this.wsFileEventDisps.forEach(x => x.dispose());
		this.wsFileEventDisps = [];
		const fe = new FileEvents(this.client, srcDir);
		this.wsFileEventDisps = fe.registerFileEvent();
	}

	async start(context: vscode.ExtensionContext, port: number, outputChannel: vscode.OutputChannel){
		const serverModule = context.asAbsolutePath(path.join('out', 'lsp-connection.js'));
		const debugOptions = { execArgv: ['--nolazy', '--inspect=6009'] };
		const serverOptions: ServerOptions = {
			run: { module: serverModule, transport: TransportKind.ipc },
			debug: {
				module: serverModule,
				transport: TransportKind.ipc,
				options: debugOptions
			}
		};
		const clientOptions: LanguageClientOptions = {
			documentSelector: [{ scheme: 'file', language: 'vb' },],
			outputChannel: outputChannel,
			initializationOptions: {
				arguments: [port.toString()] 
			}
		};
		// Create the language client and start the client.
		this.client = new LanguageClient(
			'languageServerExample',
			'Language Server Example',
			serverOptions,
			clientOptions
		);

		// Start the client. This will also launch the server
		this.client.start();

		this.waitUntilClientIsRunning();
	}

	async stop(){
		if (this.client && this.client.state === State.Running) {
			await this.client.stop();
		}
		this.wsFileEventDisps.forEach(x => x.dispose());
	}

	private async waitUntilClientIsRunning(){
		const watiTimeMs = 500;
		const countMax = 10 * 1000 / watiTimeMs;
		let waitCount = 0;
		while(true){
			if(waitCount > countMax){
				throw new Error("Timed out waiting for client ready");
			}
			waitCount++;
			await new Promise(resolve => {
				setTimeout(resolve, watiTimeMs);
			});
			if(this.client.state === State.Running){
				break;
			}
		}
	}

	async addDocuments(uris: string[]){
		const method: Hoge.RequestId = "AddDocuments";
		await this.client.sendRequest(method, {uris});
	}
	
	async diagnostics(uri: string){
		const method: Hoge.RequestId = "Diagnostic";
		await this.client.sendRequest(method, {uri});
	}
}