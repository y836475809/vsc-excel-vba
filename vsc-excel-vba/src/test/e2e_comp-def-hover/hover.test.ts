import * as assert from "assert";
import * as vscode from "vscode";
import * as helper from "../helper";

let fixtureFile : helper.FixtureFile;

suite("Extension E2E Roslyn Test Suite", () => {	
	suiteSetup(async () => {
		const port = await helper.getServerPort();
		await helper.resetServer(port);

		const filenames = ["c1.cls", "m1.bas", "m2.bas"];
		fixtureFile = new helper.FixtureFile(
			filenames
		);

		const uris = filenames.map(x => {
			return helper.getDocUri(x);
		});
		await helper.addDocuments(port, uris);

		await helper.activateExtension();
        await vscode.commands.executeCommand("vsc-excel-vba.startLanguageServer");
		await helper.sleep(500);
    });
    suiteTeardown(async () => {
    });

	test("hover class", async () => {
		const docUri = helper.getDocUri("m1.bas");
		const target = "Dim c As Class1";
		const targetPos = fixtureFile.getPosition("m1.bas", target, target.length-3);
		const expContens = [
			"```vb",
			"Class1",
			"```",
			"```xml",
			"",
			"```",
		];
		await testHover(docUri, targetPos, expContens);
	});

	test("hover class method", async () => {
		const docUri = helper.getDocUri("m1.bas");
		const target = "c1.Hello";
		const targetPos = fixtureFile.getPosition("m1.bas", target, target.length-3);
		const expContens = [
			"```vb",
			"Public Sub Hello(val1 As String)",
			"```",
			"```xml",
			"<member name=\"M:Class1.Hello(System.String)\">",
			" <summary>",
			"  メッセージ表示",
			" </summary>",
			" <param name='val1'></param>",
			" <returns></returns>",
			"</member>",
			"```",
		];
		await testHover(docUri, targetPos, expContens);
	});

	test("hover class field", async () => {
		const docUri = helper.getDocUri("m1.bas");
		const target = "c1.Name";
		const targetPos = fixtureFile.getPosition("m1.bas", target, target.length-3);
		const expContens = [
			"```vb",
			"Public Name As String",
			"```",
			"```xml",
			"<member name=\"F:Class1.Name\">",
			" <summary>",
			"  メンバ変数",
			" </summary>",
			"</member>",
			"```",
		];
		await testHover(docUri, targetPos, expContens);
	});

	test("hover module method", async () => {
		const docUri = helper.getDocUri("m1.bas");
		const target = "Sample1() ' call";
		const targetPos = fixtureFile.getPosition("m1.bas", target, 3);
		const expContens = [
			"```vb",
			"Public Sub Sample1()",
			"```",
			"```xml",
			"<member name=\"M:Module1.Sample1\">",
			" <summary>",
			"  モジュール関数",
			" </summary>",
			"</member>",
			"```",
		];
		await testHover(docUri, targetPos, expContens);
	});

	test("hover module field", async () => {
		const docUri = helper.getDocUri("m1.bas");
		const target = "buf = \"ss\"";
		const targetPos = fixtureFile.getPosition("m1.bas", target, 1);
		const expContens = [
			"```vb",
			"Private buf As String",
			"```",
			"```xml",
			"<member name=\"F:Module1.buf\">",
			" <summary>",
			"  モジュールbuf",
			" </summary>",
			"</member>",
			"```",
		];
		await testHover(docUri, targetPos, expContens);
	});

	test("hover external module method", async () => {
		const docUri = helper.getDocUri("m1.bas");
		const target = "Module2Sample1() ' call";
		const targetPos = fixtureFile.getPosition("m1.bas", target, 3);
		const expContens = [
			"```vb",
			"Public Sub Module2Sample1()",
			"```",
			"```xml",
			"<member name=\"M:Module2.Module2Sample1\">",
			" <summary>",
			"  Module2メソッド",
			" </summary>",
			"</member>",
			"```",
		];
		await testHover(docUri, targetPos, expContens);
	});

	test("hover top position", async () => {
		const docUri = helper.getDocUri("m1.bas");
		const targetPos = new vscode.Position(0, 0);
		const hover = (await vscode.commands.executeCommand(
			"vscode.executeHoverProvider",
			docUri,
			targetPos
		));
		assert.deepEqual(hover, []);
	});
});

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

