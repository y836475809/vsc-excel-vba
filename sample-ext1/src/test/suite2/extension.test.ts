import * as assert from 'assert';

// You can import and use all API from the 'vscode' module
// as well as import your extension to test it
import * as vscode from 'vscode';
import * as path from 'path';
import * as helper from '../helper';
import { LspMockServer }  from "../lsp-mock-server";

let log:string[] = [];
let mockServer: LspMockServer;

suite('Extension Test Suite', () => {	
	suiteSetup(async () => {
		mockServer = new LspMockServer();
		mockServer.callBackAddDocuments = (json: any) => {
			log.push(json.Id);
			return {};
		};
		mockServer.listen(9088);

        await vscode.commands.executeCommand("sample-ext1.startLanguageServer");
		await helper.sleep(1000);
    });
    suiteTeardown(async () => {
        await vscode.commands.executeCommand("sample-ext1.stopLanguageServer");
		mockServer.close();
    });

	test('Sample test', async () => {
		// log = [];
		await helper.activateExtension();

		// await myclinet.getComdData({});
		// try {
			// const docUri = helper.getDocUri('m1.bas');
			// const doc = await vscode.workspace.openTextDocument(docUri);
			// const editor = await vscode.window.showTextDocument(doc);
		// 	await hepler.sleep(500); // Wait for server activation
		// } catch (e) {
		// 	console.error(e);
		// }
		// const docUri = helper.getDocUri('m1.bas');
		assert.deepEqual(log, [
			"AddDocuments"
		]);
		// const docMap: any = await vscode.commands.executeCommand(
		// 	'sample-ext1.executeReverse');
		// const actStateMap = getDocStateMap(docMap);
		// assert.deepEqual(actStateMap, new Map<string, string>([
		// 	["c1.cls", "Non"],
		// 	["c2.cls", "Non"],
		// 	["collection.cls", "Non"],
		// 	["dictionary.cls", "Non"],
		// 	["m1.bas", "Non"],
		// 	["m2.bas", "Non"],
		// ]));
		// assert.deepEqual(docMap, []);
	});

	test('Sample test rename file', async () => {
		await helper.activateExtension();
		log = [];
		// await helper.sleep(500);

		const oldUri = helper.getDocUri('m2.bas');
		const newUri = helper.getDocUri("renamed_m2.bas");
		await vscode.commands.executeCommand(
			'sample-ext1.renameFiles', oldUri, newUri);
		// await vscode.workspace.fs.rename(oldUri, newUri);
		// const preff: any = await vscode.commands.executeCommand(
		// 	'sample-ext1.executeReverse');
		// const actStateMap = getDocStateMap(preff);
		// assert.deepEqual(actStateMap, new Map<string, string>([
		// 	["c1.cls", "Non"],
		// 	["c2.cls", "Non"],
		// 	["collection.cls", "Non"],
		// 	["dictionary.cls", "Non"],
		// 	["m1.bas", "Non"],
		// 	["renamed_m2.bas", "Non"],
		// ]));
		// assert.deepEqual(preff, []);
		assert.deepEqual(log, [
			"RenameDocument"
		]);
	});

	test('Sample test chnage doc', async () => {
		await helper.activateExtension();
		log = [];
		// await helper.sleep(1000);

		const docUri = helper.getDocUri('m2.bas');
		const doc = await vscode.workspace.openTextDocument(docUri);
		const editor = await vscode.window.showTextDocument(doc);
		
		assert.deepEqual(log, []);
		// assert.deepEqual(preff, new Map<string, string>([
		// 	["c1.cls", "Non"],
		// 	["c2.cls", "Non"],
		// 	["collection.cls", "Non"],
		// 	["dictionary.cls", "Non"],
		// 	["m1.bas", "Non"],
		// 	["m2.bas", "Non"],
		// ]));

		await editor.edit((editBuilder: vscode.TextEditorEdit) => {
			editBuilder.insert(new vscode.Position(0, 0), "a");
		});

		await helper.sleep(1000);

		// const postDocMap: any = await vscode.commands.executeCommand(
		// 	'sample-ext1.executeReverse');
		// const postff = getDocStateMap(postDocMap);
		// assert.deepEqual(postff, new Map<string, string>([
		// 	["c1.cls", "Non"],
		// 	["c2.cls", "Non"],
		// 	["collection.cls", "Non"],
		// 	["dictionary.cls", "Non"],
		// 	["m1.bas", "Non"],
		// 	["m2.bas", "Changed"],
		// ]));
		assert.deepEqual(log, [
			"ChangeDocument"
		]);

		await editor.edit((editBuilder: vscode.TextEditorEdit) => {
			editBuilder.delete(new vscode.Range(
				new vscode.Position(0, 0), new vscode.Position(0, 1)));
		});
		await helper.sleep(1000);
		await doc.save();

		// const saved: any = await vscode.commands.executeCommand(
		// 	'sample-ext1.executeReverse');
		// const savedStateMap = getDocStateMap(saved);
		// assert.deepEqual(savedStateMap, new Map<string, string>([
		// 	["c1.cls", "Non"],
		// 	["c2.cls", "Non"],
		// 	["collection.cls", "Non"],
		// 	["dictionary.cls", "Non"],
		// 	["m1.bas", "Non"],
		// 	["m2.bas", "Non"],
		// ]));
		assert.deepEqual(log, [
			"ChangeDocument",
			"ChangeDocument"
		]);
	});
});

function getDocStateMap(items: {[k: string]: any;}): Map<string, string>{
	const result = new Map<string, string>();
	for (const [key, value] of Object.entries(items)){
		const fname = path.basename(vscode.Uri.parse(key).fsPath);
		result.set(fname, value);
	}
	return result;
}
