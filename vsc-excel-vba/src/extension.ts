// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as vscode from 'vscode';
import {
	LanguageClient,
} from 'vscode-languageclient/node';
import { VBACommands } from './vba-commands';
import { Project } from './project';
import { Logger } from "./logger";
import { vbaRegister } from "./vba-register";
import { vbaClient } from "./vba-client";

let client: LanguageClient;
let outputChannel: vscode.OutputChannel;
let logger: Logger;

// This method is called when your extension is activated
// Your extension is activated the very first time the command is executed
export async function activate(context: vscode.ExtensionContext) {
	vscode.debug.breakpoints;

	const extName = context.extension.packageJSON.name;
	outputChannel = vscode.window.createOutputChannel(extName);
	logger = new Logger((msg: string) => {
		outputChannel.appendLine(msg);
	});

	const project = new Project("vbaproject.json");
	const vbaCommand = new VBACommands(context.asAbsolutePath("scripts"));

	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.createProject", async (args) => {
		try {
			if(project.hasProject()){
				const ret = await vscode.window.showInformationMessage(
					`${project.projectFileName} exists. Overwrite?`, "Yes", "No");
				if(ret === "No"){
					return;
				}
			}else{
				const targetFilePath = args.fsPath;		
				await project.setupConfig();
				await project.createProject(targetFilePath);
				vscode.window.showInformationMessage(`Create ${project.projectFileName}`);
				await project.readProject();
				vbaRegister(context, project, vbaCommand);
				client = vbaClient(context, project.srcDir);
				client.start();
			}
		} catch (error: unknown) {
			if(error instanceof Error){
				vscode.window.showErrorMessage(error.message);
			}else{
				vscode.window.showErrorMessage("Fail create vba project");
			}
		}
	}));
	if(project.hasProject()){
		await project.readProject();
		vbaRegister(context, project, vbaCommand);
		client = vbaClient(context, project.srcDir);
		client.start();
	}
}

// This method is called when your extension is deactivated
export async function deactivate() {
	if (!client) {
		return undefined;
	}
	return client.stop();
}
