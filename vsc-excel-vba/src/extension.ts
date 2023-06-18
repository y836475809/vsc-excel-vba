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

const registerProvider = (srcDir: string, enableLSP: boolean) => {
	const docSelector = { language: "vb" };

	dispose();

	disposables.push(vscode.languages.registerDocumentSymbolProvider(
		docSelector, new VBADocumentSymbolProvider()));

	if(enableLSP){
		vbaLSRequest.srcDir = srcDir;

		disposables.push(vscode.languages.registerSignatureHelpProvider(
			docSelector, new VBASignatureHelpProvider(vbaLSRequest), '(', ','));

		disposables.push(vscode.languages.registerDefinitionProvider(
			docSelector, new VBADefinitionProvider(vbaLSRequest)));

		disposables.push(vscode.languages.registerHoverProvider(
			docSelector, new VBAHoverProvider(vbaLSRequest)));

		disposables.push(vscode.languages.registerCompletionItemProvider(
			docSelector, new VBACompletionItemProvider(vbaLSRequest), "."));
		
		disposables.push(vscode.languages.registerReferenceProvider(
			docSelector, new VBAReferenceProvider(vbaLSRequest)));
	}
};

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
		if(await vscode.window.showInformationMessage(
			`Export?`, "Yes", "No") === "No"){
			return;
		}
		if(project.existSrcDir()){
			if(await vscode.window.showInformationMessage(
				`${project.srcDir} exists. Overwrite?`, "Yes", "No") === "No"){
				return;
			}
		}
		if(await vbaCommand.exceue(project, "export")){
			vscode.window.showInformationMessage(`Success export to ${project.srcDir}`);
		}
	}));
	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.compile", async () => {
		await vbaCommand.exceue(project, "compile");
	}));
	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.runVBASubProc", async () => {
		await vbaCommand.exceue(project, "runVBASubProc");
	}));
	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.setBreakpoints", async () => {
		await vbaCommand.exceue(project, "setBreakpoints");
	}));
	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.test.registerProvider", async () => {
		const wfs = vscode.workspace.workspaceFolders!;
		const srcDir = wfs[0].uri.fsPath;
		registerProvider(srcDir, true);
	}));

	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.createProject", async (args) => {
		try {
			if(project.hasProject()){
				const ret = await vscode.window.showInformationMessage(
					`${project.projectFileName} exists. Overwrite?`, "Yes", "No");
				if(ret === "No"){
					return;
				}
			}
			const targetFilePath = args.fsPath;		
			await project.setupConfig();
			await project.createProject(targetFilePath);
			vscode.window.showInformationMessage(`Create ${project.projectFileName}`);
		} catch (error: unknown) {
			if(error instanceof Error){
				vscode.window.showErrorMessage(error.message);
			}else{
				vscode.window.showErrorMessage("Fail create vba project");
			}
		}
	}));
	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.openSheet", async (args) => {	
		const xlsmFileName = project.projectData.targetfilename;
		await vbaCommand.openSheet(xlsmFileName, args.fsPath);
	}));

	const loadProject = async (report: (msg: string)=>void) => {
		report("Load Project");
		await project.readProject();

		registerProvider(project.srcDir, false);

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
		fileEvents.registerFileEvent();

		registerProvider(project.srcDir, true);
		
		report("Add VBA define");
		await vbaLSRequest.addVBADefines(
			loadDefinitionFiles?project.getDefinitionFileUris(context):[]);

		report("Add source");
		await vbaLSRequest.addDocuments(await project.getSrcFileUris());

		const activeDoc = vscode.window.activeTextEditor?.document;
		if(activeDoc){
			report("Diagnose active document");
			await vbaLSRequest.diagnostic(activeDoc);
		}

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
				dispose();
				fileEvents.dispose();
				
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
