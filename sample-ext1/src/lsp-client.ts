import * as path from 'path';
import * as fs from "fs";
import * as vscode from 'vscode';
import { spawn } from "child_process";
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

	async start(context: vscode.ExtensionContext, outputChannel: vscode.OutputChannel){
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
			outputChannel: outputChannel
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
		let waitCount = 0;
		while(true){
			if(waitCount > 100){
				throw new Error("Timed out waiting for client ready");
			}
			waitCount++;
			await new Promise(resolve => {
				setTimeout(resolve, 100);
			});
			if(this.client.state === State.Running){
				break;
			}
		}
	}

	async sendRequest(method: Hoge.RequestMethod, params: any){
		await this.client.sendRequest(method, params);
	}

	async shutdownServerApp(): Promise<void>{
		try {
			await this.sendRequest("Shutdown", {});
			await this.waitUntilServerApp("shutdown");
		} catch (error) {
			// 
		}
	}

	async launchServerApp(port: number, serverExeFilePath: string){
		if(await this.isReadyServerApp()){
			return;
		}
		
		if(!serverExeFilePath){
			throw new Error(`No setting value for server filepath`);
		}
		if(!fs.existsSync(serverExeFilePath)){
			throw new Error(`Not find ${serverExeFilePath}`);
		}
		const p = spawn("cmd.exe", ["/c", `${serverExeFilePath} ${port}`], { detached: true });
		p.on("error", (error)=> {
			throw error;
		});
	
		await this.waitUntilServerApp("ready");
	}

	async resetServerApp(){
		if (this.client && this.client.state === State.Running) {
			await this.client.sendRequest("reset");
		}
	}
	
	private async waitUntilServerApp(state: "ready"|"shutdown"){
		let waitCount = 0;
		while(true){
			if(waitCount > 30){
				throw new Error(`Timed out waiting for server ${state}`);
			}
			waitCount++;
			await new Promise(resolve => {
				setTimeout(resolve, 200);
			});
			try {
				await this.sendRequest("IsReady", {});
				if(state === "ready"){
					break;
				}
			} catch (error) {
				if(state === "shutdown"){
					break;
				}
			}
		}
	}
	
	private async isReadyServerApp(): Promise<boolean>{
		try {
			await this.sendRequest("IsReady", {});
			return true;
		} catch (error) {
			return false;
		}
	}
}