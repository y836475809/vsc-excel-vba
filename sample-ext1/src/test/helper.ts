import * as vscode from 'vscode';
import * as path from 'path';
import * as Mocha from 'mocha';
import * as glob from 'glob';
import * as fs from 'fs';
import { LPSRequest } from "../lsp-request";

export async function resetServer(port: number): Promise<void>{
    const lpsRequest = new LPSRequest(port);
	await lpsRequest.send({
		id: "Reset",
		filepaths: [],
        line: 0,
        chara: 0,
		text: ""
	});
	await lpsRequest.send({
		id: "IgnoreShutdown",
		filepaths: [],
        line: 0,
        chara: 0,
		text: ""
	});
};

export async function addDocuments(port: number, uris: vscode.Uri[]): Promise<void>{
    const lpsRequest = new LPSRequest(port);
    const filePaths = uris.map(uri => uri.fsPath);
    const data = {
        id: "AddDocuments",
        filepaths: filePaths,
        line: 0,
        chara: 0,
        text: ""
    } as Hoge.Command;
    await lpsRequest.send(data);
};

export class FixtureFile {
    fileMap: Map<string, string>;
    constructor(filenames: string[]){
        this.fileMap = new Map<string, string>();
        const fixtureDir = getWorkspaceFolder();
        for (const fn of filenames) {
            const fp = path.join(fixtureDir, fn);
            this.fileMap.set(fn, fs.readFileSync(fp, { encoding: "utf8"}));
        }
    }
    getText(filename: string): string {
        return this.fileMap.get(filename)!;
    }
    getPosition(filename: string, target: string, targetOffset: number): vscode.Position{
        const text = this.getText(filename);
        const index = text.indexOf(target);
        const lines = text.substring(0, index + target.length).split("\r\n");
        const lineIndex = lines.length - 1;
        const chaStart = lines[lineIndex].indexOf(target);
        return new vscode.Position(lineIndex, chaStart + targetOffset);
    }
}

export const getWorkspaceFolder = () => {
    return vscode.workspace.workspaceFolders![0].uri.fsPath;
};

export async function getServerPort(): Promise<number> {
    const config = vscode.workspace.getConfiguration("sample-ext1");
    const port: number = await config.get("serverPort")!;
    return port;
} 

export const getDocPath = (p: string) => {
    const waPath = vscode.workspace.workspaceFolders![0].uri.fsPath;
    return path.join(waPath, p);
};

export const getDocUri = (p: string) => {
    return vscode.Uri.file(getDocPath(p));
};

export function sleep(ms: number): Promise<void> {
	return new Promise(resolve => {
	  setTimeout(resolve, ms);
	});
}

export async function activateExtension() {
    const ext = vscode.extensions.getExtension('y836475809.sample-ext1')!;
    if(!ext.isActive){
        await ext.activate();
    }
}

export async function activate(docUri: vscode.Uri) {
    activateExtension();
    try {
        const doc = await vscode.workspace.openTextDocument(docUri);
        const editor = await vscode.window.showTextDocument(doc);
        await sleep(500); // Wait for server activation
    } catch (e) {
        console.error(e);
    }
}

export function run(testsRoot: string): () => Promise<void> {
	// Create the mocha test
    return (): Promise<void> =>  {
        const mocha = new Mocha({
            ui: 'tdd',
            color: true,
            timeout: 0
        });

        // const testsRoot = path.resolve(__dirname, '.');
        return new Promise((c, e) => {
            glob('**/**.test.js', { cwd: testsRoot }, (err, files) => {
                if (err) {
                    return e(err);
                }

                // Add files to the test suite
                files.forEach(f => mocha.addFile(path.resolve(testsRoot, f)));

                try {
                    // Run the mocha test
                    mocha.run(failures => {
                        if (failures > 0) {
                            e(new Error(`${failures} tests failed.`));
                        } else {
                            c();
                        }
                    });
                } catch (err) {
                    console.error(err);
                    e(err);
                }
            });
        });
    };
}