import * as vscode from 'vscode';
import * as path from 'path';
import * as Mocha from 'mocha';
import * as glob from 'glob';

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