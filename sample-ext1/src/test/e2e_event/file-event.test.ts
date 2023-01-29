import * as assert from "assert";
import * as vscode from "vscode";
import * as helper from "../helper";
import { LPSRequest } from "../../lsp-request";
import * as fs from "fs";

let lpsRequest!: LPSRequest;

async function getServerFileMap(): Promise<Map<string, string>>{
	const data = await lpsRequest.send({
		Id: "Debug:GetDocuments",
		FilePaths: [],
		Position: 0,
		Text: ""
	});
	const fileMap = new Map<string, string>();
	for (const k in data) {
		fileMap.set(k, data[k]);
	}
	return fileMap;
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

suite("Extension E2E Roslyn Test Suite", () => {	
	suiteSetup(async () => {
		const port = await helper.getServerPort();
		lpsRequest = new LPSRequest(port);
		await helper.resetServer(port);

        await vscode.commands.executeCommand("sample-ext1.startLanguageServer");
		await helper.sleep(500);
    });
    suiteTeardown(async () => {
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
		await vscode.commands.executeCommand(
			"sample-ext1.renameFiles", oldUri, newUri);

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
		await vscode.commands.executeCommand(
			"sample-ext1.renameFiles", newUri, oldUri);
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
		await helper.sleep(1000);

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
		await helper.sleep(1000);
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
		await helper.sleep(1000);

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
		await helper.sleep(1000);
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
		await helper.sleep(1000);

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
		await helper.sleep(1000);
		await doc.save();
	});

	test("delete file", async () => {
		await helper.activateExtension();

		const uri = helper.getDocUri("c1.cls");
		await vscode.commands.executeCommand(
			"sample-ext1.deleteFiles", [uri]);

		const actFileMap = await getServerFileMap();
		assert.deepEqual(actFileMap, new Map<string, string>([
			[
				helper.getDocPath("m1.bas"), 
				vbModuleCode.join("\r\n")
			]
		]));
	});
});