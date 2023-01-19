import * as assert from 'assert';

// You can import and use all API from the 'vscode' module
// as well as import your extension to test it
import * as vscode from 'vscode';
import { URI } from 'vscode-uri';
import * as path from 'path';
import * as helper from '../helper';
import * as fs from "fs";
import { LspMockServer } from "../lsp-mock-server";

let log:string[] = [];
let mockServer: LspMockServer;

suite('Extension Test Suite3', () => {
	const resMap = new Map<string, any>([
		["Completion", { items: [
			{
				DisplayText: "test1(val1)",
				CompletionText: "test1",
				Kind: "Method"
			}
		]}],
		["Definition", { items: [
			{
				FilePath: "m1.bas",
				Start: { Line: 0, Character: 0 },
				End: { Line: 0, Character: 0 },
			}
		]}],
		["Hover", { items: [
			{
				DisplayText: "test1(val1)",
				Description: "Description test1",
			}
		]}],
	]);

	suiteSetup(async () => {
		mockServer = new LspMockServer();
		mockServer.callBackAddDocuments = (json: any) => {
			log.push(json.Id);
			if(resMap.has(json.Id)){
				return resMap.get(json.Id);
			}
			return {};
		};
		const port = await helper.getServerPort();
		mockServer.listen(port);

        await vscode.commands.executeCommand("sample-ext1.startLanguageServer");
		await helper.sleep(1000);
    });
    suiteTeardown(async () => {
        await vscode.commands.executeCommand("sample-ext1.stopLanguageServer");
		mockServer.close();
    });


	test('Sample test2', async () => {
		await helper.activateExtension();
		log = [];

		const docUri = helper.getDocUri('m1.bas');
		await testCompletion(docUri, new vscode.Position(57, 7), {
			items: [
				{
					label: 'test1(val1)',
					insertText: "test1",
					kind: vscode.CompletionItemKind.Method
				}
			]
		});
		assert.deepEqual(log, [
			"Completion"
		]);
	});

	test('Sample test3', async () => {
		await helper.activateExtension();
		log = [];

		const docUri = helper.getDocUri('m1.bas');
		const exp = "```vb\ntest1(val1)\n```\n```xml\nDescription test1\n```";
		await testHover(docUri, new vscode.Position(55, 5), exp);
		assert.deepEqual(log, [
			"Hover"
		]);
	});
});

async function testCompletion(
	docUri: vscode.Uri,
	position: vscode.Position,
	expectedCompletionList: vscode.CompletionList
) {
	await helper.activate(docUri);

	// Executing the command `vscode.executeCompletionItemProvider` to simulate triggering completion
	const actualCompletionList = (await vscode.commands.executeCommand(
		'vscode.executeCompletionItemProvider',
		docUri,
		position
	)) as vscode.CompletionList;

	const aItems = actualCompletionList.items;
	const actualCompletionItems = actualCompletionList.items.filter(x => {
		return x.kind === vscode.CompletionItemKind.Method;
	});
	assert.ok(actualCompletionItems.length === 1);
	expectedCompletionList.items.forEach((expectedItem, i) => {
		const actualItem = actualCompletionItems[i];
		assert.equal(actualItem.label, expectedItem.label);
		assert.equal(actualItem.insertText, expectedItem.insertText);
		assert.equal(actualItem.kind, expectedItem.kind);
	});
}

async function testHover(
	docUri: vscode.Uri,
	position: vscode.Position,
	expectedContents: string
) {
	await helper.activate(docUri);

	// Executing the command `vscode.executeCompletionItemProvider` to simulate triggering completion
	const hover = (await vscode.commands.executeCommand(
		'vscode.executeHoverProvider',
		docUri,
		position
	)) as vscode.Hover[];
	const contents = <vscode.MarkdownString[]>hover[0].contents;
	const actValue = contents[0].value;
	assert.equal(actValue, expectedContents);
}
