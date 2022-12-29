// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as path from 'path';
import * as vscode from 'vscode';

import {
	LanguageClient,
	LanguageClientOptions,
	ServerOptions,
	TransportKind
} from 'vscode-languageclient/node';
import { TreeDataProvider } from './treeDataProvider';

let client: LanguageClient;

// This method is called when your extension is activated
// Your extension is activated the very first time the command is executed
export async function activate(context: vscode.ExtensionContext) {
	const wf = vscode.workspace.workspaceFolders;
	const rootPath = (wf && (wf.length > 0)) ? wf[0].uri.fsPath : undefined;
	
	const treeDataProvider = new TreeDataProvider();
	vscode.window.registerTreeDataProvider('testView', treeDataProvider);
	vscode.commands.registerCommand('testView.addEntry', () => console.log("testView.addEntry"));
	vscode.commands.registerCommand('testView.myCommand', (arg: string) => {
		console.log("testView.myCommand=", arg);

	});
	
	// const uris = await vscode.workspace.findFiles("*.{bas,cls}");
	// const fsPaths = uris.map(uri => uri.fsPath);

	// const editorConfig: any = await vscode.workspace.getConfiguration().get('editor.quickSuggestions');
	// if(editorConfig["other"] !== "off"){
	// 	editorConfig["other"] = "on";
	// 	await vscode.workspace.getConfiguration().update(
	// 		'editor.quickSuggestions',
	// 		editorConfig,
	// 		vscode.ConfigurationTarget.Workspace
	// 	);
	// }

	vscode.window.onDidChangeActiveTextEditor((e) => {
		// e.
	});

	let serverModule = context.asAbsolutePath(path.join('out', 'server.js'));
	let debugOptions = { execArgv: ['--nolazy', '--inspect=6009'] };
	let serverOptions: ServerOptions = {
		run: { module: serverModule, transport: TransportKind.ipc },
		debug: {
			module: serverModule,
			transport: TransportKind.ipc,
			options: debugOptions
		}
	};
	let clientOptions: LanguageClientOptions = {
		// Register the server for plain text documents
		documentSelector: [
			// { scheme: 'file', language: 'plaintext' },
			// { scheme: "untitled", language: "plaintext" },
			{ scheme: 'file', language: 'vb' },
			// { scheme: 'file', language: 'visual basic' }
		],
	};

	// Create the language client and start the client.
	client = new LanguageClient(
		'languageServerExample',
		'Language Server Example',
		serverOptions,
		clientOptions
	);

	// Start the client. This will also launch the server
	client.start();
}

// This method is called when your extension is deactivated
export function deactivate(): Thenable<void> | undefined {
	if (!client) {
		return undefined;
	}
	return client.stop();
}
