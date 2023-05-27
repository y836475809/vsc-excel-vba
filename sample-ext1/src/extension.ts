// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as path from 'path';
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

function setupOutline(context: vscode.ExtensionContext) {
	if(outlineDisp){
		outlineDisp.dispose();
	}
	outlineDisp = vscode.languages.registerDocumentSymbolProvider(
		{ language: "vb" }, new VbaDocumentSymbolProvider());
	context.subscriptions.push(outlineDisp);
}

function setServerStartEnable(enable: boolean){
	vscode.commands.executeCommand("setContext", "sample-ext1.start.enable", enable);
}

// This method is called when your extension is activated
// Your extension is activated the very first time the command is executed
export async function activate(context: vscode.ExtensionContext) {
	setServerStartEnable(true);
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
	let serverAppPort = await config.get("sample-ext1.serverPort") as number;
	let loadDefinitionFiles = await config.get("sample-ext1.loadDefinitionFiles");
	let useLSPServer = await config.get("sample-ext1.useLSPServer") as boolean;
	let autoLaunchServerApp = await config.get("sample-ext1.autoLaunchServerApp") as boolean;
	let serverExeFilePath = await config.get("sample-ext1.serverExeFilePath") as string;

	vscode.workspace.onDidChangeConfiguration(async event => {
		if(!event.affectsConfiguration("sample-ext1")){
			return;
		}
		const config = vscode.workspace.getConfiguration();
		serverAppPort = await config.get("sample-ext1.serverPort") as number;
		loadDefinitionFiles = await config.get("sample-ext1.loadDefinitionFiles");
		useLSPServer = await config.get("sample-ext1.useLSPServer") as boolean;
		autoLaunchServerApp = await config.get("sample-ext1.autoLaunchServerApp") as boolean;
		serverExeFilePath = await config.get("sample-ext1.serverExeFilePath") as string;
	});

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
		await vbaCommand.exceue(project, "import");
	}));
	context.subscriptions.push(vscode.commands.registerCommand("sample-ext1.export", async () => {
		await vbaCommand.exceue(project, "export");
	}));
	context.subscriptions.push(vscode.commands.registerCommand("sample-ext1.compile", async () => {
		await vbaCommand.exceue(project, "compile");
	}));

	context.subscriptions.push(vscode.commands.registerCommand("sample-ext1.startLanguageServer", async () => {
		await lspClient.start(context, outputChannel);	
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

	context.subscriptions.push(vscode.commands.registerCommand("sample-ext1.runVBASubProc", async (args) => {
		const editor = vscode.window.activeTextEditor;
		if(!editor){
			return;
		}
		const sel = editor.selection;
		const uri = editor.document.uri;
		const symbols = await vscode.commands.executeCommand<vscode.DocumentSymbol[]>(
			"vscode.executeDocumentSymbolProvider", uri);
		if(!symbols.length){
			return;
		}
		const targetSymbols = symbols[0].children.filter(x => {
			return x.kind === vscode.SymbolKind.Function 
				&& x.range.start.line <= sel.start.line 
				&& sel.end.line <= x.range.end.line;
		});
		if(!targetSymbols.length){
			return;
		}
		const moduleName = path.parse(uri.fsPath).name;;
		const symName = targetSymbols[0].detail;
		const mt = symName.match(/Sub\s+(.+)\(\s*\)/);
		if(mt && mt.length > 1){
			const procName = mt[1];
			const xlsmFileName = project.projectData.targetfilename;
			await vbaCommand.runVBASubProc(xlsmFileName, `${moduleName}.${procName}`);
		}
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

		report("stop ServerAppr");
		await lspClient.stop();	

		if(autoLaunchServerApp){
			report("shutdownServerApp");
			await lspClient.shutdownServerApp(serverAppPort);
			report("launchServerApp");
			await lspClient.launchServerApp(serverAppPort, serverExeFilePath);
		}

		report("Initialize ServerAppr");
		await lspClient.start(context, outputChannel);
		
		report("resetServerApp");
		await lspClient.resetServerApp();

		lspClient.registerFileEvents(project.srcDir);

		const uris1 = await project.getSrcFileUris();
		const uris2 = loadDefinitionFiles?project.getDefinitionFileUris(context):[];
		const uris = uris1.concat(uris2);
		if(uris.length > 0){
			report("Send source uris to ServerAppr");
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
			setServerStartEnable(false);
			try {	
				await loadProject((msg) => {
					logger.info(msg);
					progress.report({ message: msg });
				});
				if(useLSPServer){
					await startServer((msg) => {
						logger.info(msg);
						progress.report({ message: msg });
					});
				}
				vscode.window.showInformationMessage("Success");
			} catch (error) {
				let errorMsg = "Fail start";
				if(error instanceof Error){
					errorMsg = `${error.message}`;
				}
				logger.info(`Fail start, ${errorMsg}`);
				vscode.window.showErrorMessage(`Fail start\n${errorMsg}\nPlease restart again`);
			}
			setServerStartEnable(true);
		});
	}));

	context.subscriptions.push(vscode.commands.registerCommand("sample-ext1.stop", async () => {
		if(outlineDisp){
			outlineDisp.dispose();
		}
		await lspClient.stop();
		await lspClient.shutdownServerApp(serverAppPort);
	}));
}

// This method is called when your extension is deactivated
export async function deactivate() {
	// TODO
	const config = vscode.workspace.getConfiguration();
	let serverAppPort = await config.get("sample-ext1.serverPort") as number;
	await lspClient.shutdownServerApp(serverAppPort);
	await lspClient.stop();
}
