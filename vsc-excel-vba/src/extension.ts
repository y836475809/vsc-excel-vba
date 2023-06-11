// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as vscode from 'vscode';
import { MyTreeItem, TreeDataProvider } from './treeDataProvider';
import { VBACommands } from './vba-commands';
import { Project } from './project';
import { Logger } from "./logger";
import { FileEvents } from "./file-events";
import { VBALSLaunch } from "./vba-ls-launch";
import { VBADocumentSymbolProvider } from "./vba-document-symbol-provider";
import { VBASignatureHelpProvider } from "./vba-signaturehelp-provider";
import { VBADefinitionProvider } from "./vba-definition-provider";
import { VBAHoverProvider } from "./vba-hover-provider";
import { VBACompletionItemProvider } from "./vba-completionitem-provider";
import { VBAReferenceProvider } from "./vba-reference-provider";
import { VBALSRequest } from './vba-ls-request';

let outputChannel: vscode.OutputChannel;
let logger: Logger;
let fileEvents: FileEvents;
let vbaLSRequest: VBALSRequest;
let vbaLSLaunch: VBALSLaunch;
let statusBarItem: vscode.StatusBarItem;
let disposables: vscode.Disposable[] = [];

function dispose() {
	disposables.forEach(x=>x.dispose());
	disposables = [];
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

	const config = vscode.workspace.getConfiguration();
	let loadDefinitionFiles = await config.get("vsc-excel-vba.loadDefinitionFiles");
	let enableLSP = await config.get("vsc-excel-vba.enableLSP") as boolean;
	let vbaLanguageServerPort = await config.get("vsc-excel-vba.VBALanguageServerPort") as number;
	let enableAutoLaunchVBALanguageServer = await config.get("vsc-excel-vba.enableAutoLaunchVBALanguageServer") as boolean;
	let vbaLanguageServerPath = await config.get("vsc-excel-vba.VBALanguageServerPath") as string;
	let enableVBACompileAfterImport = await config.get("vsc-excel-vba.enableVBACompileAfterImport") as boolean;

	vbaLSRequest = new VBALSRequest(vbaLanguageServerPort);
	fileEvents = new FileEvents(vbaLSRequest);
	vbaLSLaunch = new VBALSLaunch(vbaLanguageServerPort);
	const project = new Project("vbaproject.json");
	const vbaCommand = new VBACommands(context.asAbsolutePath("scripts"));

    statusBarItem = vscode.window.createStatusBarItem(
		vscode.StatusBarAlignment.Left, 1);
	statusBarItem.text = `Run ${extName}`;
	statusBarItem.command = `${extName}.toggle`;
	statusBarItem.show();

	vscode.workspace.onDidChangeConfiguration(async event => {
		if(!event.affectsConfiguration("vsc-excel-vba")){
			return;
		}
		const config = vscode.workspace.getConfiguration();
		loadDefinitionFiles = await config.get("vsc-excel-vba.loadDefinitionFiles");
		enableLSP = await config.get("vsc-excel-vba.enableLSP") as boolean;
		vbaLanguageServerPort = await config.get("vsc-excel-vba.VBALanguageServerPort") as number;
		enableAutoLaunchVBALanguageServer = await config.get("vsc-excel-vba.enableAutoLaunchVBALanguageServer") as boolean;
		vbaLanguageServerPath = await config.get("vsc-excel-vba.VBALanguageServerPath") as string;
		enableVBACompileAfterImport = await config.get("vsc-excel-vba.enableVBACompileAfterImport") as boolean;
	});

	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.toggle", async () => {
		if(statusBarItem.text === `Run ${extName}`){
			await vscode.commands.executeCommand("vsc-excel-vba.start");
		}else{
			await vscode.commands.executeCommand("vsc-excel-vba.stop");
		}
	}));

	vscode.window.registerTreeDataProvider("testView", new TreeDataProvider());
	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.runViewItem", async (args: MyTreeItem) => {
		await vbaCommand.exceue(project, args.vbaCommand);
	}));
	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.gotoVBA", async () => {
		await vbaCommand.exceue(project, "gotoVBA");
	}));
	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.gotoVSCode", async () => {
		await vbaCommand.exceue(project, "gotoVSCode");
	}));
	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.import", async () => {
		if(enableVBACompileAfterImport){
			await vbaCommand.exceue(project, "importAndCompile");
		}else{
			await vbaCommand.exceue(project, "import");
		}
	}));
	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.export", async () => {
		await vbaCommand.exceue(project, "export");
	}));
	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.compile", async () => {
		await vbaCommand.exceue(project, "compile");
	}));
	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.runVBASubProc", async () => {
		await vbaCommand.exceue(project, "runVBASubProc");
	}));
	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.test.registerProvider", async () => {
		registerProvider();
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
	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.openSheet", async (args) => {	
		const xlsmFileName = project.projectData.targetfilename;
		await vbaCommand.openSheet(xlsmFileName, args.fsPath);
	}));

	const registerProvider = () => {
		dispose();

		const docSelector = { language: "vb" };

		disposables.push(vscode.languages.registerDocumentSymbolProvider(
			docSelector, new VBADocumentSymbolProvider()));

		if(enableLSP){
			disposables.push(vscode.languages.registerSignatureHelpProvider(
				docSelector, new VBASignatureHelpProvider(vbaLanguageServerPort), '(', ','));

			disposables.push(vscode.languages.registerDefinitionProvider(
				docSelector, new VBADefinitionProvider(vbaLanguageServerPort)));

			disposables.push(vscode.languages.registerHoverProvider(
				docSelector, new VBAHoverProvider(vbaLanguageServerPort)));

			disposables.push(vscode.languages.registerCompletionItemProvider(
				docSelector, new VBACompletionItemProvider(vbaLanguageServerPort), "."));
			
			disposables.push(vscode.languages.registerReferenceProvider(
				docSelector, new VBAReferenceProvider(vbaLanguageServerPort)));
		}
	};

	const loadProject = async (report: (msg: string)=>void) => {
		report("Load Project");
		await project.readProject();

		registerProvider();

		report("Loaded project successfully");
	};

	const startServer = async (report: (msg: string)=>void) => {
		report("Start server");

		if(enableAutoLaunchVBALanguageServer){
			report("Shutdown VBALanguageServer");
			await vbaLSLaunch.shutdown();
			report("Start VBALanguageServer");
			await vbaLSLaunch.start(vbaLanguageServerPort, vbaLanguageServerPath);
		}else{
			report("Reset VBALanguageServer");
			await vbaLSLaunch.reset();
		}

		report("Register FileEvent");
		fileEvents.registerFileEvent(project.srcDir);

		const uris1 = await project.getSrcFileUris();
		const uris2 = loadDefinitionFiles?project.getDefinitionFileUris(context):[];
		const uris = uris1.concat(uris2);
		if(uris.length > 0){
			report("Add source");
			await vbaLSRequest.addDocuments(uris);
		}

		const activeDoc = vscode.window.activeTextEditor?.document;
		if(activeDoc){
			const activePath = activeDoc.uri.fsPath;
			if(activePath.endsWith(".bas") || activePath.endsWith(".cls")){
				report("Diagnose active document");
				await vbaLSRequest.diagnostic(activeDoc);
			}
		}

		registerProvider();

		report("Server started successfully");
	};

	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.start", async () => {
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

	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.stop", async () => {
		dispose();
		fileEvents.dispose();

		await vbaLSLaunch.shutdown();
		statusBarItem.text = `Run ${extName}`;
	}));
}

// This method is called when your extension is deactivated
export async function deactivate() {
	dispose();
	fileEvents.dispose();

	await vbaLSLaunch.shutdown();
}
