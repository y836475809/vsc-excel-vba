import * as vscode from 'vscode';
import { LPSRequest } from './lsp-request';

export class VBAReferenceProvider implements vscode.ReferenceProvider {
    lpsRequest: LPSRequest;
    constructor(port: number){
        this.lpsRequest = new LPSRequest(port);
    }
	async provideReferences(document: vscode.TextDocument, position: vscode.Position, context: vscode.ReferenceContext, token: vscode.CancellationToken)
	: Promise<vscode.Location[]> {
		try {
			const uri = document.uri;
			const fp = uri.fsPath;
			const data = {
				id: "References",
				filepaths: [fp],
				line: position.line,
				chara: position.character,
				text: ""
			} as Hoge.RequestParam;
			const items = await this.lpsRequest.send(data) as Hoge.ReferencesItem[];
			const locs = items.map(x => {
				return new vscode.Location(
					uri,
					new vscode.Range(
						x.start.line, x.start.character,
						x.end.line, x.end.character
					)  
				);
			});
			return locs;
		} catch (error) {
			if(error instanceof Error){
				vscode.window.showErrorMessage(error.message);
			}else{
				vscode.window.showErrorMessage(String(error));
			}
			return [];
		}
	}
}
