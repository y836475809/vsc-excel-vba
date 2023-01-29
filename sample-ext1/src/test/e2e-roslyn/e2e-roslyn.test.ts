import * as assert from "assert";
import * as vscode from "vscode";
import * as helper from "../helper";
import { LPSRequest } from "../../lsp-request";
import * as fs from "fs";

let lpsRequest!: LPSRequest;
let fixtureData : helper.FixtureData;

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

function getPosition(filename: string, target: string, targetOffset: number): vscode.Position{
	const text = fixtureData.getText(filename);
	const index = text.indexOf(target);
	const lines = text.substring(0, index + target.length).split("\r\n");
	const lineIndex = lines.length - 1;
	const chaStart = lines[lineIndex].indexOf(target);
	return new vscode.Position(lineIndex, chaStart + targetOffset);
}

type FileMap = Map<string, string>;
function assertFileMap(actFileMap: FileMap, expFileMap: FileMap){
	assert.equal(actFileMap.size, expFileMap.size);
	const actKeys = Array.from(actFileMap.keys()).sort();
	const expKeys = Array.from(expFileMap.keys()).sort();
	assert.deepEqual(actKeys, expKeys);
	actFileMap.forEach((v, k) => {
		const expv = expFileMap.get(k)!;
		const actLines = v.split("\r\n");
		const expLines = expv.split("\r\n");
		for (let index = 0; index < actLines.length; index++) {
			const act = actLines[index];
			const exp = expLines[index];
			assert.equal(act, exp, `index=${index}, key=${k}`);
		}
	});
}


suite("Extension E2E Roslyn Test Suite", () => {	
	suiteSetup(async () => {
		const port = await helper.getServerPort();
		lpsRequest = new LPSRequest(port);
		fixtureData = new helper.FixtureData();

        await vscode.commands.executeCommand("sample-ext1.startLanguageServer");
		await helper.sleep(500);
    });
    suiteTeardown(async () => {
		const renamed = helper.getDocPath("renamed_m2.bas");
		if(fs.existsSync(renamed)){
			const renamedUri = helper.getDocUri("renamed_m2.bas");
			const orgUri = helper.getDocUri("m2.bas");
			helper.getDocPath("renamed_m2.bas");
			await vscode.workspace.fs.rename(renamedUri, orgUri);
		}
    });

	// test("init", async () => {
	// 	await helper.activateExtension();

	// 	const actFileMap = await getServerFileMap();
	// 	const expFileMap = fixtureData.getVbFileMap();
	// 	assertFileMap(actFileMap, expFileMap);
	// });

	// test("rename file", async () => {
	// 	await helper.activateExtension();

	// 	const oldUri = helper.getDocUri("m2.bas");
	// 	const newUri = helper.getDocUri("renamed_m2.bas");
	// 	await vscode.workspace.fs.rename(oldUri, newUri);
	// 	await vscode.commands.executeCommand(
	// 		"sample-ext1.renameFiles", oldUri, newUri);

	// 	const actFileMap = await getServerFileMap();
	// 	const expFileMap = fixtureData.getVbFileMap();
	// 	fixtureData.rename(expFileMap, "m2.bas", "renamed_m2.bas");

	// 	const renamedPath = helper.getDocPath("renamed_m2.bas");
	// 	// const oldText = expFileMap.get(renamedPath)!;
	// 	// expFileMap.set(renamedPath, oldText.replace("Module m2", "Module renamed_m2"));
	// 	// assert.deepEqual(actFileMap, expFileMap);
	// 	assertFileMap(actFileMap, expFileMap);

	// 	await vscode.workspace.fs.rename(newUri, oldUri);
	// 	await vscode.commands.executeCommand(
	// 		"sample-ext1.renameFiles", newUri, oldUri);
	// });

	test("change content", async () => {
		await helper.activateExtension();

		const actFileMap = await getServerFileMap();
		const expFileMap = fixtureData.getVbFileMap();
		// assert.deepEqual(actFileMap, expFileMap);
		assertFileMap(actFileMap, expFileMap);

		const docUri = helper.getDocUri("m2.bas");
		const doc = await vscode.workspace.openTextDocument(docUri);
		const editor = await vscode.window.showTextDocument(doc);
		await editor.edit((editBuilder: vscode.TextEditorEdit) => {
			editBuilder.insert(new vscode.Position(1, 0), "a");
		});
		await helper.sleep(1000);

		const actFileMapChanged = await getServerFileMap();
		const expFileMapChanged = fixtureData.getVbFileMap();
		// const orgText = fixtureData.getText("m2.bas");
		fixtureData.update(expFileMapChanged, "m2.bas", 
			["Module m2", `asample`, "End Module"].join("\r\n"));
		assertFileMap(actFileMapChanged, expFileMapChanged);

		await editor.edit((editBuilder: vscode.TextEditorEdit) => {
			editBuilder.delete(new vscode.Range(
				new vscode.Position(1, 0), new vscode.Position(1, 1)));
		});
		await helper.sleep(1000);

		await editor.edit((editBuilder: vscode.TextEditorEdit) => {
			editBuilder.insert(new vscode.Position(0, 21), "a");
		});
		await helper.sleep(1000);
		const actFileMapChanged2 = await getServerFileMap();
		const expFileMapChanged2 = fixtureData.getVbFileMap();
		// const orgText2 = fixtureData.getText("m2.bas");
		fixtureData.update(expFileMapChanged2, "m2.bas", 
			["Module am2", `sample`, "End Module"].join("\r\n"));
		assertFileMap(actFileMapChanged2, expFileMapChanged2);

		await editor.edit((editBuilder: vscode.TextEditorEdit) => {
			editBuilder.delete(new vscode.Range(
				new vscode.Position(0, 21), new vscode.Position(0, 22)));
		});

		await helper.sleep(1000);
		await doc.save();

		const actFileMapSaved = await getServerFileMap();
		// assert.deepEqual(actFileMapSaved, expFileMap);
		assertFileMap(actFileMapSaved, expFileMap);
	});

	// test("delete file", async () => {
	// 	await helper.activateExtension();

	// 	const uri = helper.getDocUri("c2.cls");
	// 	await vscode.commands.executeCommand(
	// 		"sample-ext1.deleteFiles", [uri]);

	// 	const actFileMap = await getServerFileMap();
	// 	const expFileMap = fixtureData.getVbFileMap();
	// 	fixtureData.delete(expFileMap, "c2.cls");
	// 	// assert.deepEqual(actFileMap, expFileMap);
	// 	assertFileMap(actFileMap, expFileMap);
	// });

	// test("completion class", async () => {
	// 	await helper.activateExtension();

	// 	const docUri = helper.getDocUri("m1.bas");
	// 	const target = "testList.Item";
	// 	const targetPos = getPosition("m1.bas", target, target.length - "Item".length);
	// 	await testCompletion(docUri, targetPos, {
	// 		items: [
	// 			{
	// 				label: "Public Count As Long",
	// 				insertText: "Count",
	// 				kind: vscode.CompletionItemKind.Field
	// 			},
	// 			{
	// 				label: "Public Property Item(index As Integer) As Object",
	// 				insertText: "Item",
	// 				kind: vscode.CompletionItemKind.Property
	// 			},
	// 			{
	// 				label: "Public Property Item(key As String) As Object",
	// 				insertText: "Item",
	// 				kind: vscode.CompletionItemKind.Property
	// 			},
	// 			{
	// 				label: "Public Sub Add(item As Object)",
	// 				insertText: "Add",
	// 				kind: vscode.CompletionItemKind.Method
	// 			},
	// 			{
	// 				label: "Public Sub Add(item As Object, key As String)",
	// 				insertText: "Add",
	// 				kind: vscode.CompletionItemKind.Method
	// 			},
	// 			{
	// 				label: "Public Sub Remove(index As Long)",
	// 				insertText: "Remove",
	// 				kind: vscode.CompletionItemKind.Method
	// 			},
	// 			{
	// 				label: "Public Sub Remove(key As String)",
	// 				insertText: "Remove",
	// 				kind: vscode.CompletionItemKind.Method
	// 			},
	// 		]
	// 	});
	// });

	// test("completion module", async () => {
	// 	await helper.activateExtension();

	// 	const docUri = helper.getDocUri("m1.bas");
	// 	const target = "' completion position";
	// 	const pos = getPosition("m1.bas", target, target.length);
	// 	const targetPos = new vscode.Position(pos.line + 1, 0);
	// 	await testCompletion(docUri, targetPos, {
	// 		items: [
	// 			{
	// 				label: "Collection",
	// 				insertText: "Collection",
	// 				kind: vscode.CompletionItemKind.Class
	// 			},
	// 			{
	// 				label: "Dictionary",
	// 				insertText: "Dictionary",
	// 				kind: vscode.CompletionItemKind.Class
	// 			},
	// 			{
	// 				label: "Person",
	// 				insertText: "Person",
	// 				kind: vscode.CompletionItemKind.Class
	// 			},
	// 			{
	// 				label: "Private buf As String",
	// 				insertText: "buf",
	// 				kind: vscode.CompletionItemKind.Field
	// 			},
	// 			{
	// 				label: "Public Sub Sample1()",
	// 				insertText: "Sample1",
	// 				kind: vscode.CompletionItemKind.Method
	// 			},
	// 			{
	// 				label: "Public Sub Sample2()",
	// 				insertText: "Sample2",
	// 				kind: vscode.CompletionItemKind.Method
	// 			},
	// 			{
	// 				label: "Public Sub Sample3()",
	// 				insertText: "Sample3",
	// 				kind: vscode.CompletionItemKind.Method
	// 			},
	// 			{
	// 				label: "Public Sub Sample4()",
	// 				insertText: "Sample4",
	// 				kind: vscode.CompletionItemKind.Method
	// 			},
	// 			{
	// 				label: "Public Sub call1()",
	// 				insertText: "call1",
	// 				kind: vscode.CompletionItemKind.Method
	// 			},
	// 			{
	// 				label: "p2",
	// 				insertText: "p2",
	// 				kind: vscode.CompletionItemKind.Variable
	// 			},
	// 		]
	// 	});
	// });

	// test("definition class", async () => {
	// 	await helper.activateExtension();

	// 	const docUri = helper.getDocUri("m1.bas");
	// 	const target = "Dim p2 As New Person";
	// 	const targetPos = getPosition("m1.bas", target, target.length-3);

	// 	const defTarget = "Public Class Person";
	// 	const defStartPos = getPosition("c1.cls", defTarget, defTarget.length - "Person".length);
	// 	const defEndPos = getPosition("c1.cls", defTarget, defTarget.length);
	// 	await testDefinition(docUri, targetPos, 
	// 		[
	// 			new vscode.Location(
	// 				helper.getDocUri("c1.cls"), new vscode.Range(defStartPos, defEndPos))
	// 		]
	// 	);
	// });

	// test("definition module field", async () => {
	// 	await helper.activateExtension();

	// 	const docUri = helper.getDocUri("m1.bas");
	// 	const target = "buf = \"ss\"";
	// 	const targetPos = getPosition("m1.bas", target, 1);

	// 	const defTarget = "Private buf As String";
	// 	const defStartPos = getPosition("m1.bas", defTarget, "Private ".length);
	// 	const defEndPos = getPosition("m1.bas", defTarget, "Private buf".length);
	// 	await testDefinition(docUri, targetPos, 
	// 		[
	// 			new vscode.Location(
	// 				helper.getDocUri("m1.bas"), new vscode.Range(defStartPos, defEndPos))
	// 		]
	// 	);
	// });

	// test("definition class method", async () => {
	// 	await helper.activateExtension();

	// 	const docUri = helper.getDocUri("m1.bas");
	// 	const target = "p2.SayHello";
	// 	const targetPos = getPosition("m1.bas", target, target.length - 3);

	// 	const defTarget = "Public Sub SayHello(val1, val2)";
	// 	const defStartPos = getPosition("c1.cls", defTarget, "Public Sub ".length);
	// 	const defEndPos = getPosition("c1.cls", defTarget, "Public Sub SayHello".length);
	// 	await testDefinition(docUri, targetPos, 
	// 		[
	// 			new vscode.Location(
	// 				helper.getDocUri("c1.cls"), new vscode.Range(defStartPos, defEndPos))
	// 		]
	// 	);
	// });

	// test("hover class", async () => {
	// 	await helper.activateExtension();

	// 	const docUri = helper.getDocUri("m1.bas");
	// 	const target = "Dim p2 As New Person";
	// 	const targetPos = getPosition("m1.bas", target, target.length-3);
	// 	const expContens = [
	// 		"```vb",
	// 		"Person",
	// 		"```",
	// 		"```xml",
	// 		"<member name=\"T:Person\">",
	// 		" <summary>",
	// 		"  個人",
	// 		" </summary>",
	// 		"</member>",
	// 		"```",
	// 	];
	// 	await testHover(docUri, targetPos, expContens);
	// });

	// test("hover class method", async () => {
	// 	await helper.activateExtension();

	// 	const docUri = helper.getDocUri("m1.bas");
	// 	const target = "p2.SayHello";
	// 	const targetPos = getPosition("m1.bas", target, target.length-3);
	// 	const expContens = [
	// 		"```vb",
	// 		"Public Sub SayHello(val1 As Object, val2 As Object)",
	// 		"```",
	// 		"```xml",
	// 		"<member name=\"M:Person.SayHello(System.Object,System.Object)\">",
	// 		" <summary>",
	// 		"  テストメッセージ",
	// 		" </summary>",
	// 		" <param name='val1'></param>",
	// 		" <param name='val2'></param>",
	// 		" <returns></returns>",
	// 		"</member>",
	// 		"```",
	// 	];
	// 	await testHover(docUri, targetPos, expContens);
	// });
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

	const actualCompletionItems = actualCompletionList.items.filter(x => {
		return x.kind !== vscode.CompletionItemKind.Snippet;
	});
	actualCompletionItems.sort((a, b): number => {
		return a.label < b.label?-1:1;
	});
	assert.equal(actualCompletionItems.length, expectedCompletionList.items.length);
	expectedCompletionList.items.forEach((expectedItem, i) => {
		const actualItem = actualCompletionItems[i];
		assert.equal(actualItem.label, expectedItem.label);
		assert.equal(actualItem.insertText, expectedItem.insertText);
		assert.equal(actualItem.kind, expectedItem.kind);
	});
}

async function testDefinition(
	docUri: vscode.Uri,
	position: vscode.Position,
	expectedLocationList: vscode.Location[]
) {
	await helper.activate(docUri);
	const actualLocationList = (await vscode.commands.executeCommand(
		"vscode.executeDefinitionProvider",
		docUri,
		position
	)) as vscode.Location[];

	assert.equal(actualLocationList.length, expectedLocationList.length);
	expectedLocationList.forEach((expectedItem, i) => {
		const actualItem = actualLocationList[i];
		assert.equal(actualItem.uri.toString(), expectedItem.uri.toString());
		assert.ok(actualItem.range.isEqual(expectedItem.range));
	});
}

async function testHover(
	docUri: vscode.Uri,
	position: vscode.Position,
	expectedContents: string[]
) {
	await helper.activate(docUri);
	const hover = (await vscode.commands.executeCommand(
		"vscode.executeHoverProvider",
		docUri,
		position
	)) as vscode.Hover[];
	const contents = <vscode.MarkdownString[]>hover[0].contents;
	const actValues = contents[0].value.split("\n");
	expectedContents.forEach((expContent, i) => {
		const actVal = actValues[i];
		assert.equal(actVal, expContent);
	});	
}

