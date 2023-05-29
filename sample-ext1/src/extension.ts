// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as path from 'path';
import * as fs from "fs";
import * as vscode from 'vscode';
import { MyTreeItem, TreeDataProvider } from './treeDataProvider';
import { VBACommands } from './vba-commands';
import { Project } from './project';
import { VbaDocumentSymbolProvider } from "./vba-documentsymbolprovider";
import { Logger } from "./logger";
import { LSPClient } from "./lsp-client";

let outlineDisp: vscode.Disposable;
let outputChannel: vscode.OutputChannel;
let logger: Logger;
let lspClient: LSPClient;
let statusBarItem: vscode.StatusBarItem;

function setupOutline(context: vscode.ExtensionContext) {
	if(outlineDisp){
		outlineDisp.dispose();
	}
	outlineDisp = vscode.languages.registerDocumentSymbolProvider(
		{ language: "vb" }, new VbaDocumentSymbolProvider());
	context.subscriptions.push(outlineDisp);
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

	lspClient = new LSPClient();
	const project = new Project("project.json");
	const vbaCommand = new VBACommands(context.asAbsolutePath("scripts"));

	const config = vscode.workspace.getConfiguration();
	let loadDefinitionFiles = await config.get("sample-ext1.loadDefinitionFiles");
	let enableLSP = await config.get("sample-ext1.enableLSP") as boolean;
	let vbaLanguageServerPort = await config.get("sample-ext1.VBALanguageServerPort") as number;
	let enableAutoLaunchVBALanguageServer = await config.get("sample-ext1.enableAutoLaunchVBALanguageServer") as boolean;
	let vbaLanguageServerPath = await config.get("sample-ext1.VBALanguageServerPath") as string;
	let enableVBACompileAfterImport = await config.get("sample-ext1.enableVBACompileAfterImport") as boolean;

    statusBarItem = vscode.window.createStatusBarItem(
		vscode.StatusBarAlignment.Left, 1);
	statusBarItem.text = `Run ${extName}`;
	statusBarItem.command = `${extName}.toggle`;
	statusBarItem.show();

	vscode.workspace.onDidChangeConfiguration(async event => {
		if(!event.affectsConfiguration("sample-ext1")){
			return;
		}
		const config = vscode.workspace.getConfiguration();
		loadDefinitionFiles = await config.get("sample-ext1.loadDefinitionFiles");
		enableLSP = await config.get("sample-ext1.enableLSP") as boolean;
		vbaLanguageServerPort = await config.get("sample-ext1.VBALanguageServerPort") as number;
		enableAutoLaunchVBALanguageServer = await config.get("sample-ext1.enableAutoLaunchVBALanguageServer") as boolean;
		vbaLanguageServerPath = await config.get("sample-ext1.VBALanguageServerPath") as string;
		enableVBACompileAfterImport = await config.get("sample-ext1.enableVBACompileAfterImport") as boolean;
	});

	context.subscriptions.push(vscode.commands.registerCommand("sample-ext1.toggle", async () => {
		if(statusBarItem.text === `Run ${extName}`){
			await vscode.commands.executeCommand("sample-ext1.start");
		}else{
			await vscode.commands.executeCommand("sample-ext1.stop");
		}
	}));

	vscode.window.registerTreeDataProvider("testView", new TreeDataProvider());
	context.subscriptions.push(vscode.commands.registerCommand("sample-ext1.runViewItem", async (args: MyTreeItem) => {
		await vbaCommand.exceue(project, args.vbaCommand);
	}));
	context.subscriptions.push(vscode.commands.registerCommand("sample-ext1.gotoVBA", async () => {
		await vbaCommand.exceue(project, "gotoVBA");
	}));
	context.subscriptions.push(vscode.commands.registerCommand("sample-ext1.gotoVSCode", async () => {
		await vbaCommand.exceue(project, "gotoVSCode");
	}));
	context.subscriptions.push(vscode.commands.registerCommand("sample-ext1.import", async () => {
		if(enableVBACompileAfterImport){
			await vbaCommand.exceue(project, "importAndCompile");
		}else{
			await vbaCommand.exceue(project, "import");
		}
	}));
	context.subscriptions.push(vscode.commands.registerCommand("sample-ext1.export", async () => {
		await vbaCommand.exceue(project, "export");
	}));
	context.subscriptions.push(vscode.commands.registerCommand("sample-ext1.compile", async () => {
		await vbaCommand.exceue(project, "compile");
	}));
	context.subscriptions.push(vscode.commands.registerCommand("sample-ext1.runVBASubProc", async () => {
		await vbaCommand.exceue(project, "runVBASubProc");
	}));
	context.subscriptions.push(vscode.commands.registerCommand("sample-ext1.startLanguageServer", async () => {
		await lspClient.start(context, vbaLanguageServerPort, outputChannel);	
	}));

	context.subscriptions.push(vscode.commands.registerCommand("testView.myCommand", async (args) => {
		try {
			if(project.hasProject()){
				throw new Error(
					`Not find ${project.projectFileName}. create project`);
			}
			const targetFilePath = args.fsPath;		
			await project.setupConfig();
			await project.createProject(targetFilePath);
			vscode.window.showInformationMessage(`Create ${project.projectFileName}`);
		} catch (error: unknown) {
			if(error instanceof Error){
				vscode.window.showErrorMessage(error.message);
			}else{
				vscode.window.showErrorMessage("Fail create project");
			}
		}
	}));
	context.subscriptions.push(vscode.commands.registerCommand("sample-ext1.openSheet", async (args) => {	
		const xlsmFileName = project.projectData.targetfilename;
		await vbaCommand.openSheet(xlsmFileName, args.fsPath);
	}));

	const loadProject = async (report: (msg: string)=>void) => {
		report("Load Project");
		await project.readProject();

		report("Setup Outline");
		setupOutline(context);

		report("Loaded project successfully");
	};

	const startServer = async (report: (msg: string)=>void) => {
		report("Start server");

		report("stop VBALanguageServer");
		await lspClient.stop();	

		report("Initialize VBALanguageServer");
		await lspClient.start(context, vbaLanguageServerPort, outputChannel);

		if(enableAutoLaunchVBALanguageServer){
			report("shutdownVBALanguageServer");
			await lspClient.shutdownVBALanguageServer();
			report("launchVBALanguageServer");
			await lspClient.launchVBALanguageServer(vbaLanguageServerPort, vbaLanguageServerPath);
		}else{
			report("resetVBALanguageServer");
			await lspClient.resetVBALanguageServer();
		}

		lspClient.registerFileEvents(project.srcDir);

		const uris1 = await project.getSrcFileUris();
		const uris2 = loadDefinitionFiles?project.getDefinitionFileUris(context):[];
		const uris = uris1.concat(uris2);
		if(uris.length > 0){
			report("Send source uris to VBALanguageServer");
			const method: Hoge.RequestMethod = "createFiles";
			lspClient.sendRequest(method, {uris});
		}

		const activeUri = vscode.window.activeTextEditor?.document.uri;
		if(activeUri){
			if(activeUri.fsPath.endsWith(".bas") || activeUri.fsPath.endsWith(".cls")){
				report("Diagnose active document");
				const method: Hoge.RequestMethod = "diagnostics";
				lspClient.sendRequest(method, {uri:activeUri.toString()});
			}
		}

		report("Server started successfully");
	};

	context.subscriptions.push(vscode.commands.registerCommand("sample-ext1.start", async () => {
		const options: vscode.ProgressOptions = {
			location: vscode.ProgressLocation.Notification,
			title: "Server status"
		};
		vscode.window.withProgress(options, async progress => {
			try {	
				await loadProject((msg) => {
					logger.info(msg);
					progress.report({ message: msg });
				});
				if(enableLSP){
					await startServer((msg) => {
						logger.info(msg);
						progress.report({ message: msg });
					});
				}
				vscode.window.showInformationMessage("Success");
				statusBarItem.text = `Stop ${extName}`;
			} catch (error) {
				await lspClient.stop();	

				let errorMsg = "Fail start";
				if(error instanceof Error){
					errorMsg = `${error.message}`;
				}
				logger.info(`Fail start, ${errorMsg}`);
				vscode.window.showErrorMessage(`Fail start\n${errorMsg}\nPlease restart again`);
				statusBarItem.text = `Run ${extName}`;
			}
		});
	}));

	context.subscriptions.push(vscode.commands.registerCommand("sample-ext1.stop", async () => {
		if(outlineDisp){
			outlineDisp.dispose();
		}
		await lspClient.shutdownVBALanguageServer();
		await lspClient.stop();
		statusBarItem.text = `Run ${extName}`;
	}));
}

// This method is called when your extension is deactivated
export async function deactivate() {
	await lspClient.shutdownVBALanguageServer();
	await lspClient.stop();
}
