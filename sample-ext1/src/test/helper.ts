import * as vscode from 'vscode';
import * as path from 'path';
import * as Mocha from 'mocha';
import * as glob from 'glob';
import * as fs from 'fs';

export class FixtureData {
    fileMap: Map<string, string>;
    pathMap: Map<string, string>;;
    constructor(){
        this.pathMap = new Map<string, string>();
        this.fileMap = new Map<string, string>();
        const fixtureDir = getWorkspaceFolder();
        const fsPaths = [
            path.join(fixtureDir, ".vscode", "collection.cls"),
            path.join(fixtureDir, ".vscode", "dictionary.cls"),
            path.join(fixtureDir, "c1.cls"),
            path.join(fixtureDir, "c2.cls"),
            path.join(fixtureDir, "m1.bas"),
            path.join(fixtureDir, "m2.bas"),
        ];
        for (const fp of fsPaths) {
            this.pathMap.set(path.basename(fp), fp);
            this.fileMap.set(fp, fs.readFileSync(fp, { encoding: "utf8"}));
        }
    }

    getText(filename: string): string {
        const fp = this.pathMap.get(filename)!;
        return this.fileMap.get(fp)!;
    }

    getFileMap(): Map<string, string> { 
        const cloneMap = new Map<string, string>(
            JSON.parse(JSON.stringify(Array.from(this.fileMap)))
        );
        return cloneMap;
    }

    rename(fileMap: Map<string, string>, oldFilename: string, newFilename: string){
        const oldFp = this.pathMap.get(oldFilename)!;
        const text = fileMap.get(oldFp)!;
		fileMap.delete(oldFp);
        const newFp = path.join(path.dirname(oldFp), newFilename);
		fileMap.set(newFp, text);
    }

    update(fileMap: Map<string, string>, filename:string, text: string){
        const fp = this.pathMap.get(filename)!;
		fileMap.set(fp, text);
    }

    delete(fileMap: Map<string, string>, filename:string){
        const fp = this.pathMap.get(filename)!;
		fileMap.delete(fp);
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