import * as assert from "assert";
import * as vscode from "vscode";
import * as helper from "../helper";

let fixtureFile : helper.FixtureFile;

suite("Extension E2E Completion Test Suite", () => {	
	suiteSetup(async () => {
		fixtureFile = new helper.FixtureFile(
			["c1.cls", "m1.bas", "m2.bas"]
		);

		const port = await helper.getServerPort();
		await helper.resetServer(port);

		const uris = [
			helper.getDocUri("m1.bas"),
			helper.getDocUri("m2.bas"),
			helper.getDocUri("c1.cls"),
		];
		await helper.addDocuments(port, uris);

		await helper.activateExtension();
		await vscode.commands.executeCommand("vsc-excel-vba.test.registerProvider");

		await helper.sleep(500);
    });
    suiteTeardown(async () => {
    });

	test("completion class", async () => {
		const docUri = helper.getDocUri("m1.bas");
		const target = "c1.";
		const targetPos = fixtureFile.getPosition("m1.bas", target, target.length);
		await testCompletion(docUri, targetPos, {
			items: [
				{
					label: "Name",
					detail: "Public Name As String",
					insertText: "Name",
					kind: vscode.CompletionItemKind.Field
				},
				{
					label: "Age",
					detail: "Public Age As Long",
					insertText: "Age",
					kind: vscode.CompletionItemKind.Field
				},
				{
					label: "Prop1",
					detail: "Public Property Prop1(index As Long) As Long",
					insertText: "Prop1",
					kind: vscode.CompletionItemKind.Property
				},
				{
					label: "Hello",
					detail: "Public Sub Hello(val1 As String)",
					insertText: "Hello",
					kind: vscode.CompletionItemKind.Method
				},
				{
					label: "GetHello",
					detail: "Public Function GetHello(val As String) As String",
					insertText: "GetHello",
					kind: vscode.CompletionItemKind.Method
				}
			]
		});
	});

	test("completion module", async () => {
		const docUri = helper.getDocUri("m1.bas");
		const target = "' completion position";
		const pos = fixtureFile.getPosition("m1.bas", target, target.length);
		const targetPos = new vscode.Position(pos.line + 1, 0);
		await testCompletion(docUri, targetPos, {
			items: [
				{
					label: "Class1",
					detail: "Class1",
					insertText: "Class1",
					kind: vscode.CompletionItemKind.Class
				},
				{
					label: "buf",
					detail: "Private buf As String",
					insertText: "buf",
					kind: vscode.CompletionItemKind.Field
				},
				{
					label: "Sample1",
					detail: "Public Sub Sample1()",
					insertText: "Sample1",
					kind: vscode.CompletionItemKind.Method
				},
				{
					label: "Sample2",
					detail: "Private Function Sample2() As String",
					insertText: "Sample2",
					kind: vscode.CompletionItemKind.Method
				},
				{
					label: "Sample3",
					detail: "Private Sub Sample3()",
					insertText: "Sample3",
					kind: vscode.CompletionItemKind.Method
				},
				{
					label: "Module2Sample1",
					detail: "public Sub Module2Sample1()",
					insertText: "Module2Sample1",
					kind: vscode.CompletionItemKind.Method
				},
				{
					label: "call1",
					detail: "Public Sub call1()",
					insertText: "call1",
					kind: vscode.CompletionItemKind.Method
				},
				{
					label: "c",
					detail: "c",
					insertText: "c",
					kind: vscode.CompletionItemKind.Variable
				},
				{
					label: "c1",
					detail: "c1",
					insertText: "c1",
					kind: vscode.CompletionItemKind.Variable
				},
			]
		});
	});

	test("completion top position", async () => {
		const docUri = helper.getDocUri("m1.bas");
		const targetPos = new vscode.Position(0, 0);
		await testCompletion(docUri, targetPos, {items: []});
	});
});

async function testCompletion(
	docUri: vscode.Uri,
	position: vscode.Position,
	expectedCompletionList: vscode.CompletionList
) {
	await helper.activate(docUri);
	const actualCompletionList = (await vscode.commands.executeCommand(
		"vscode.executeCompletionItemProvider",
		docUri,
		position
	)) as vscode.CompletionList;

	const sotrFunc = (a:any, b:any): number => {
		return a.label < b.label?-1:1;
	};

	const actualCompletionItems = actualCompletionList.items.filter(x => {
		return x.kind !== vscode.CompletionItemKind.Snippet
			&& x.kind !== vscode.CompletionItemKind.Text
			&& x.kind !== vscode.CompletionItemKind.Keyword;
	}).sort(sotrFunc);
	const expectedCompletionItems = expectedCompletionList.items.sort(sotrFunc);

	assert.equal(actualCompletionItems.length, expectedCompletionItems.length);
	actualCompletionItems.forEach((expectedItem, i) => {
		const actualItem = actualCompletionItems[i];
		assert.equal(actualItem.label, expectedItem.label);
		assert.equal(actualItem.insertText, expectedItem.insertText);
		assert.equal(actualItem.kind, expectedItem.kind);
	});
}
