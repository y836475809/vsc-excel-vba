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
import { LPSRequest } from "./lsp-request";
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

	async sendRequest(method: Hoge.RequestMethod, params: any){
		await this.client.sendRequest(method, params);
	}

	async shutdownServerApp(port: number): Promise<void>{
		const req = new LPSRequest(port);
		const data = {
			id: "Shutdown",
			filepaths: [],
			line: 0,
			chara: 0,
			text: ""
		} as Hoge.Command;
		try {
			await req.send(data);
			await this.waitUntilServerApp(port, "shutdown");
		} catch (error) {
			// 
		}
	}

	async launchServerApp(port: number, serverExeFilePath: string){
		if(await this.isReadyServerApp(port)){
			return;
		}
		
		const prop = "sample-ext1.serverExeFilePath";
		if(!serverExeFilePath){
			throw new Error(`${prop} is not set`);
		}
		if(!fs.existsSync(serverExeFilePath)){
			throw new Error(`${prop}, Not find: ${serverExeFilePath}`);
		}
		const p = spawn("cmd.exe", ["/c", `${serverExeFilePath} ${port}`], { detached: true });
		p.on("error", (error)=> {
			throw error;
		});
	
		await this.waitUntilServerApp(port, "ready");
	}

	async resetServerApp(){
		if (this.client && this.client.state === State.Running) {
			const method: Hoge.RequestMethod = "reset";
			await this.client.sendRequest(method);
		}
	}
	
	private async waitUntilServerApp(port: number, state: "ready"|"shutdown"){
		const req = new LPSRequest(port);
		let waitCount = 0;
		while(true){
			if(waitCount > 30){
				throw new Error(`Timed out waiting for server ${state}`);
			}
			waitCount++;
			await new Promise(resolve => {
				setTimeout(resolve, 200);
			});
			const data = {
				id: "IsReady",
				filepaths: [],
				line: 0,
				chara: 0,
				text: ""
			} as Hoge.Command;
			try {
				await req.send(data);
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
	
	private async isReadyServerApp(port: number): Promise<boolean>{
		const req = new LPSRequest(port);
		const data = {
			id: "IsReady",
			filepaths: [],
			line: 0,
			chara: 0,
			text: ""
		} as Hoge.Command;
		try {
			await req.send(data);
			return true;
		} catch (error) {
			return false;
		}
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
}