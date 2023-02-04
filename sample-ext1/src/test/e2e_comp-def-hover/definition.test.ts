import * as assert from "assert";
import * as vscode from "vscode";
import * as helper from "../helper";

let fixtureFile : helper.FixtureFile;

function getPosition(filename: string, target: string, targetOffset: number): vscode.Position{
	const text = fixtureFile.getText(filename);
	const index = text.indexOf(target);
	const lines = text.substring(0, index + target.length).split("\r\n");
	const lineIndex = lines.length - 1;
	const chaStart = lines[lineIndex].indexOf(target);
	return new vscode.Position(lineIndex, chaStart + targetOffset);
}

suite("Extension E2E Definition Test Suite", () => {	
	suiteSetup(async () => {
		const port = await helper.getServerPort();
		await helper.resetServer(port);

		fixtureFile = new helper.FixtureFile(
			["c1.cls", "m1.bas", "m2.bas"]
		);

        await vscode.commands.executeCommand("sample-ext1.startLanguageServer");
		await helper.sleep(500);
    });
    suiteTeardown(async () => {
    });

	test("definition class", async () => {
		await helper.activateExtension();

		const docUri = helper.getDocUri("m1.bas");
		const target = "Dim c As Class1";
		const targetPos = getPosition("m1.bas", target, target.length-3);

		const defTarget = "Attribute VB_Name = \"";
		// const defStartPos = getPosition("c1.cls", defTarget, defTarget.length);
		// const defEndPos = getPosition("c1.cls", defTarget, defTarget.length + "Class1".length);
		const zeroPos = new vscode.Position(0, 0);
		await testDefinition(docUri, targetPos, 
			[
				new vscode.Location(
					helper.getDocUri("c1.cls"), new vscode.Range(zeroPos, zeroPos))
			]
		);
	});

	test("definition class constructor", async () => {
		await helper.activateExtension();

		const docUri = helper.getDocUri("m1.bas");
		const target = "Dim c1 As New Class1";
		const targetPos = getPosition("m1.bas", target, target.length-3);

		const defTarget = "Attribute VB_Name = \"";
		// const defStartPos = getPosition("c1.cls", defTarget, defTarget.length);
		// const defEndPos = getPosition("c1.cls", defTarget, defTarget.length + "Class1".length);
		const zeroPos = new vscode.Position(0, 0);
		await testDefinition(docUri, targetPos, 
			[
				new vscode.Location(
					helper.getDocUri("c1.cls"), new vscode.Range(zeroPos, zeroPos))
			]
		);
	});

	test("definition module field", async () => {
		await helper.activateExtension();

		const docUri = helper.getDocUri("m1.bas");
		const target = "buf = \"ss\"";
		const targetPos = getPosition("m1.bas", target, 1);

		const defTarget = "Private buf As String";
		const defStartPos = getPosition("m1.bas", defTarget, "Private ".length);
		const defEndPos = getPosition("m1.bas", defTarget, "Private buf".length);
		await testDefinition(docUri, targetPos, 
			[
				new vscode.Location(
					helper.getDocUri("m1.bas"), new vscode.Range(defStartPos, defEndPos))
			]
		);
	});

	test("definition module method", async () => {
		await helper.activateExtension();

		const docUri = helper.getDocUri("m1.bas");
		const target = "Sample1() ' call";
		const targetPos = getPosition("m1.bas", target, 1);

		const defTarget = "Public Sub Sample1()";
		const defStartPos = getPosition("m1.bas", defTarget, "Public Sub ".length);
		const defEndPos = getPosition("m1.bas", defTarget, "Public Sub Sample1".length);
		await testDefinition(docUri, targetPos, 
			[
				new vscode.Location(
					helper.getDocUri("m1.bas"), new vscode.Range(defStartPos, defEndPos))
			]
		);
	});

	test("definition external module method", async () => {
		await helper.activateExtension();

		const docUri = helper.getDocUri("m1.bas");
		const target = "Module2Sample1() ' call";
		const targetPos = getPosition("m1.bas", target, 1);

		const defTarget = "Public Sub Module2Sample1()";
		const defStartPos = getPosition("m2.bas", defTarget, "Public Sub ".length);
		const defEndPos = getPosition("m2.bas", defTarget, "Public Sub Module2Sample1".length);
		await testDefinition(docUri, targetPos, 
			[
				new vscode.Location(
					helper.getDocUri("m2.bas"), new vscode.Range(defStartPos, defEndPos))
			]
		);
	});

	test("definition class method", async () => {
		await helper.activateExtension();

		const docUri = helper.getDocUri("m1.bas");
		const target = "c1.Hello";
		const targetPos = getPosition("m1.bas", target, target.length - 3);

		const defTarget = "Public Sub Hello(val1 As String)";
		const defStartPos = getPosition("c1.cls", defTarget, "Public Sub ".length);
		const defEndPos = getPosition("c1.cls", defTarget, "Public Sub Hello".length);
		await testDefinition(docUri, targetPos, 
			[
				new vscode.Location(
					helper.getDocUri("c1.cls"), new vscode.Range(defStartPos, defEndPos))
			]
		);
	});

	test("definition class field", async () => {
		await helper.activateExtension();

		const docUri = helper.getDocUri("m1.bas");
		const target = "c1.Name";
		const targetPos = getPosition("m1.bas", target, target.length - 3);

		const defTarget = "Public Name As String";
		const defStartPos = getPosition("c1.cls", defTarget, "Public ".length);
		const defEndPos = getPosition("c1.cls", defTarget, "Public Name".length);
		await testDefinition(docUri, targetPos, 
			[
				new vscode.Location(
					helper.getDocUri("c1.cls"), new vscode.Range(defStartPos, defEndPos))
			]
		);
	});

	test("definition top position", async () => {
		await helper.activateExtension();

		const docUri = helper.getDocUri("m1.bas");
		const targetPos = new vscode.Position(0, 0);
		const actualLocationList = (await vscode.commands.executeCommand(
			"vscode.executeDefinitionProvider",
			docUri,
			targetPos
		));
		assert.deepEqual(actualLocationList, []);
	});
});

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
