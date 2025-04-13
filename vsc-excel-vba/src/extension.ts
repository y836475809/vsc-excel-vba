// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as path from "path";
import * as vscode from 'vscode';
import {
	LanguageClient,
	LanguageClientOptions,
	ServerOptions,
	TransportKind,
} from 'vscode-languageclient/node';
import { VBACommands } from './vba-commands';
import { Project } from './project';
import { Logger } from "./logger";
import { VBACreateFile } from "./vba-create-file";
import { VBADocumentSymbolProvider } from "./vba-document-symbol-provider";
import { SheetTreeDataProvider } from "./sheet-treedata-provider";

let client: LanguageClient;
let outputChannel: vscode.OutputChannel;
let logger: Logger;
let statusBarItem: vscode.StatusBarItem;
let disposables: vscode.Disposable[] = [];
let sheetTDProvider: SheetTreeDataProvider;

const registerProviderSideBar = () => {
	const docSelector = { language: "vb" };

	disposables.push(vscode.languages.registerDocumentSymbolProvider(
		docSelector, new VBADocumentSymbolProvider()));

	sheetTDProvider = new SheetTreeDataProvider("sheetView");
	disposables.push(sheetTDProvider);
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

	const project = new Project("vbaproject.json");
	const vbaCommand = new VBACommands(context.asAbsolutePath("scripts"));

	await project.readProject();
	registerProviderSideBar();

    // statusBarItem = vscode.window.createStatusBarItem(
	// 	vscode.StatusBarAlignment.Left, 1);
	// statusBarItem.text = `Run ${extName}`;
	// statusBarItem.command = `${extName}.toggle`;
	// statusBarItem.show();

	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.gotoVBA", async () => {
		await vbaCommand.exceue(project, "gotoVBA");
	}));
	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.gotoVSCode", async () => {
		await vbaCommand.exceue(project, "gotoVSCode");
	}));
	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.import", async () => {
		await vbaCommand.exceue(project, "importAndCompile");
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
	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.runVBAProc", async () => {
		await vbaCommand.exceue(project, "runVBAProc");
	}));
	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.setBreakpoints", async () => {
		await vbaCommand.exceue(project, "setBreakpoints");
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
		const xlsmFileName = project.projectData.excelfilename;
		await vbaCommand.openSheet(xlsmFileName, args.sheetName);
	}));
	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.getSheetNames", async (args) => {	
		const xlsmFileName = project.projectData.excelfilename;
		const sheetNames = await vbaCommand.getSheetNames(xlsmFileName);
		sheetTDProvider.refresh(sheetNames);
	}));
	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.newClassFile", async (uri: vscode.Uri) => {	
		const vbacf = new VBACreateFile();
		await vbacf.create(uri, "class");
	}));
	context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.newModuleFile", async (uri: vscode.Uri) => {	
		const vbacf = new VBACreateFile();
		await vbacf.create(uri, "module");
	}));

	const srcDirName = path.basename(project.srcDir);
	const config = vscode.workspace.getConfiguration();
	const lspFilename = config.get("vsc-excel-vba.LSFilename") as string;
	const exeFilePath = path.join(context.asAbsolutePath("bin"), lspFilename);
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
			fileEvents: vscode.workspace.createFileSystemWatcher(`${project.srcDir}/*.{bas,cls}`)
		},
	};
	client = new LanguageClient(
		"VBALanguageServer",
		"VBA Language Server",
		serverOptions,
		clientOptions
	);
	client.start();
}

// This method is called when your extension is deactivated
export async function deactivate() {
	if (!client) {
		return undefined;
	}
	return client.stop();
}
