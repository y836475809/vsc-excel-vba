import * as assert from "assert";
import * as vscode from "vscode";
import * as helper from "../helper";
import { LPSRequest } from "../../lsp-request";
import * as fs from "fs";
import { URI } from 'vscode-uri';

let lpsRequest!: LPSRequest;

async function getServerFileMap(): Promise<Map<string, string>>{
	const data = await lpsRequest.send({
		id: "Debug:GetDocuments",
		filepaths: [],
		line: 0,
		chara: 0,
		text: ""
	});
	const json = JSON.parse(data);
	const fileMap = new Map<string, string>();
	for (const k in json) {
		fileMap.set(k, json[k]);
	}
	return fileMap;
}

async function renameFile(oldUri: vscode.Uri, newUri: vscode.Uri){
	const oldFsPath = oldUri.fsPath;
	const newFsPath = newUri.fsPath;
	const data = {
		id: "RenameDocument",
		filepaths: [oldFsPath, newFsPath],
		line: 0,
		chara: 0,
		text: ""
	} as Hoge.Command;  
	await lpsRequest.send(data);
}

async function deleteFile(uri: vscode.Uri){
	const data = {
		id: "DeleteDocuments",
		filepaths: [uri.fsPath],
		line: 0,
		chara: 0,
		text: ""
	} as Hoge.Command;  
	await lpsRequest.send(data);
}

const vbClassCode = [
	"Public Class Class1", 
	"Public class_value As String",
	"End Class"	
];

const vbModuleCode = [
	"Module Module1", 
	"Private module_value As String",
	"End Module"
];

let disps: vscode.Disposable[]  = [];

suite("Extension E2E Roslyn Test Suite", () => {	
	suiteSetup(async () => {
		const port = await helper.getServerPort();
		lpsRequest = new LPSRequest(port);
		await helper.resetServer(port);

		const uris = [
			helper.getDocUri("m1.bas"),
			helper.getDocUri("c1.cls"),
		];
		await helper.addDocuments(port, uris);

		const disp = vscode.workspace.onDidChangeTextDocument(async (e: vscode.TextDocumentChangeEvent) =>{
			const fsPath = e.document.uri.fsPath;
            const data = {
                id: "ChangeDocument",
                filepaths: [fsPath],
                position: 0,
                line: 0,
                chara: 0,
                text: e.document.getText()
            } as Hoge.Command;
            await lpsRequest.send(data);
		});
		disps.push(disp);

		await helper.sleep(500);
    });
    suiteTeardown(async () => {
		disps.forEach(x => {
			x.dispose();
		});

		const renamed = helper.getDocPath("re_m1.bas");
		if(fs.existsSync(renamed)){
			const renamedUri = helper.getDocUri("re_m1.bas");
			const orgUri = helper.getDocUri("m1.bas");
			await vscode.workspace.fs.rename(renamedUri, orgUri);
		}
    });

	test("init", async () => {
		await helper.activateExtension();

		const actFileMap = await getServerFileMap();
		assert.deepEqual(actFileMap, new Map<string, string>([
			[
				helper.getDocPath("c1.cls"), 
				vbClassCode.join("\r\n")
			], [
				helper.getDocPath("m1.bas"), 
				vbModuleCode.join("\r\n")
			]
		]));
	});

	test("rename file", async () => {
		await helper.activateExtension();

		const oldUri = helper.getDocUri("m1.bas");
		const newUri = helper.getDocUri("re_m1.bas");
		await vscode.workspace.fs.rename(oldUri, newUri);
		await renameFile(oldUri, newUri);

		const actFileMap = await getServerFileMap();
		assert.deepEqual(actFileMap, new Map<string, string>([
			[
				helper.getDocPath("c1.cls"), 
				vbClassCode.join("\r\n")
			], [
				helper.getDocPath("re_m1.bas"), 
				vbModuleCode.join("\r\n")
			]
		]));

		await vscode.workspace.fs.rename(newUri, oldUri);
		await renameFile(newUri, oldUri);
	});

	test("change content", async () => {
		await helper.activateExtension();

		const actFileMap = await getServerFileMap();
		assert.deepEqual(actFileMap, new Map<string, string>([
			[
				helper.getDocPath("c1.cls"), 
				vbClassCode.join("\r\n")
			], [
				helper.getDocPath("m1.bas"), 
				vbModuleCode.join("\r\n")
			]
		]));

		const docUri = helper.getDocUri("m1.bas");
		const doc = await vscode.workspace.openTextDocument(docUri);
		const editor = await vscode.window.showTextDocument(doc);
		await editor.edit((editBuilder: vscode.TextEditorEdit) => {
			editBuilder.insert(new vscode.Position(1, 0), "a");
		});
		await helper.sleep(1500);

		const actFileMapChanged = await getServerFileMap();
		assert.deepEqual(actFileMapChanged, new Map<string, string>([
			[
				helper.getDocPath("c1.cls"), 
				vbClassCode.join("\r\n")
			], [
				helper.getDocPath("m1.bas"), 
				[
					"Module Module1", 
					"aPrivate module_value As String",
					"End Module"
				].join("\r\n")
			]
		]));

		await editor.edit((editBuilder: vscode.TextEditorEdit) => {
			editBuilder.delete(new vscode.Range(
				new vscode.Position(1, 0), new vscode.Position(1, 1)));
		});
		await helper.sleep(1500);
		await doc.save();
	});

	test("change module name", async () => {
		await helper.activateExtension();

		const actFileMap = await getServerFileMap();
		assert.deepEqual(actFileMap, new Map<string, string>([
			[
				helper.getDocPath("c1.cls"), 
				vbClassCode.join("\r\n")
			], [
				helper.getDocPath("m1.bas"), 
				vbModuleCode.join("\r\n")
			]
		]));

		const attVBname = "Attribute VB_Name = \"Module1\"";
		const reAttVBname = "Attribute VB_Name = \"ReModule1\"";
		const docUri = helper.getDocUri("m1.bas");
		const doc = await vscode.workspace.openTextDocument(docUri);
		const editor = await vscode.window.showTextDocument(doc);
		await editor.edit((editBuilder: vscode.TextEditorEdit) => {
			editBuilder.replace(new vscode.Range(
				new vscode.Position(0, 0), new vscode.Position(0, attVBname.length)), 
				reAttVBname);
		});
		await helper.sleep(1500);

		const actFileMapChanged = await getServerFileMap();
		assert.deepEqual(actFileMapChanged, new Map<string, string>([
			[
				helper.getDocPath("c1.cls"), 
				vbClassCode.join("\r\n")
			], [
				helper.getDocPath("m1.bas"), 
				[
					"Module ReModule1", 
					"Private module_value As String",
					"End Module"
				].join("\r\n")
			]
		]));

		await editor.edit((editBuilder: vscode.TextEditorEdit) => {
			editBuilder.replace(new vscode.Range(
				new vscode.Position(0, 0), new vscode.Position(0, reAttVBname.length)), 
				attVBname);
		});
		await helper.sleep(1500);
		await doc.save();
	});

	test("change class name", async () => {
		await helper.activateExtension();

		const actFileMap = await getServerFileMap();
		assert.deepEqual(actFileMap, new Map<string, string>([
			[
				helper.getDocPath("c1.cls"), 
				vbClassCode.join("\r\n")
			], [
				helper.getDocPath("m1.bas"), 
				vbModuleCode.join("\r\n")
			]
		]));

		const attVBname = "Attribute VB_Name = \"Class1\"";
		const reAttVBname = "Attribute VB_Name = \"ReClass1\"";
		const docUri = helper.getDocUri("c1.cls");
		const doc = await vscode.workspace.openTextDocument(docUri);
		const editor = await vscode.window.showTextDocument(doc);
		await editor.edit((editBuilder: vscode.TextEditorEdit) => {
			editBuilder.replace(new vscode.Range(
				new vscode.Position(4, 0), new vscode.Position(4, attVBname.length)), 
				reAttVBname);
		});
		await helper.sleep(1500);

		const actFileMapChanged = await getServerFileMap();
		assert.deepEqual(actFileMapChanged, new Map<string, string>([
			[
				helper.getDocPath("c1.cls"), 
				[
					"Public Class ReClass1", 
					"Public class_value As String",
					"End Class"	
				].join("\r\n")
			], [
				helper.getDocPath("m1.bas"), 
				vbModuleCode.join("\r\n")
			]
		]));

		await editor.edit((editBuilder: vscode.TextEditorEdit) => {
			editBuilder.replace(new vscode.Range(
				new vscode.Position(4, 0), new vscode.Position(4, reAttVBname.length)), 
				attVBname);
		});
		await helper.sleep(1500);
		await doc.save();
	});

	test("delete file", async () => {
		await helper.activateExtension();

		const uri = helper.getDocUri("c1.cls");
		await deleteFile(uri);

		const actFileMap = await getServerFileMap();
		assert.deepEqual(actFileMap, new Map<string, string>([
			[
				helper.getDocPath("m1.bas"), 
				vbModuleCode.join("\r\n")
			]
		]));
	});
});