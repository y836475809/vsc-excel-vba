// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as vscode from 'vscode';
import * as fs from "fs";
import { LanguageClient, State } from 'vscode-languageclient/node';
import { Commands } from './commands';
import { Project } from './project';
import { Logger } from "./logger";
import { register } from "./register";
import { createClient } from "./client";

let client: LanguageClient | undefined;
let outputChannel: vscode.OutputChannel;
let logger: Logger;

async function getServerConfig(): Promise<[boolean, string]>{
	const section = "vsc-excel-vba.languageServer";
	const config = vscode.workspace.getConfiguration();
	const serverEnable = config.get<boolean>(`${section}.enable`, true);
	const serverPath = config.get<string>(`${section}.path`, "");
	if(!serverEnable){
		return [false, ""];
	}
	if(fs.existsSync(serverPath)){
		return [true, serverPath];
	}
	
	const ret = await vscode.window.showInformationMessage(
		`VBA language server path does not exists. Open setting?`, "Yes", "No");
	if(ret === "Yes"){
		await vscode.commands.executeCommand(
			"workbench.action.openSettings", `${section}.path`);
	}
	return [false, ""];
}


// This method is called when your extension is activated
// Your extension is activated the very first time the command is executed
export async function activate(context: vscode.ExtensionContext) {
	vscode.debug.breakpoints;

	const extName = context.extension.packageJSON.name;
	outputChannel = vscode.window.createOutputChannel(extName);
	logger = new Logger((msg: string) => {
		outputChannel.appendLine(msg);
	});

	const project = new Project();
	const commands = new Commands(context.asAbsolutePath("scripts"));

	vscode.workspace.onDidChangeConfiguration(async (e) => {
		const section = "vsc-excel-vba.languageServer";
		if (e.affectsConfiguration(`${section}.path`) || e.affectsConfiguration(`${section}.enable`)) {		
			const config = vscode.workspace.getConfiguration();
			const serverEnable = config.get<boolean>(`${section}.enable`, true);
			const serverPath = config.get<string>(`${section}.path`, "");
			if(!serverEnable){
				if (client?.state === State.Running) {
					await client.stop();
				}
				return;
			}
			if(!fs.existsSync(serverPath)){
				await vscode.window.showInformationMessage(
					"VBA language server path does not exists.");
				return;
			}
			if (client?.state === State.Running) {
				await client.stop();
			}
			client = createClient(context, serverPath, project.srcDir);
			client.start();
		}
	});

	const setup = async () => {
		const [serverEnable, serverPath] = await getServerConfig();
		await project.readProject();
		register(context, project, commands);
		if(serverEnable){
			client = createClient(context, serverPath, project.srcDir);
			client.start();
		}
	};

	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.createProject", async (args) => {
		try {
			const targetFilePath = args.fsPath;		
			await project.createProject(targetFilePath);
			await setup();
		} catch (error: unknown) {
			if(error instanceof Error){
				vscode.window.showErrorMessage(error.message);
			}else{
				vscode.window.showErrorMessage("Unknown error");
			}
		}
	}));
	if(project.hasProject()){
		await setup();
	}
}

// This method is called when your extension is deactivated
export async function deactivate() {
	if (!client) {
		return undefined;
	}
	return client.stop();
}
