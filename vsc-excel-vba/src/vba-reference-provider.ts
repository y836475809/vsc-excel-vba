import * as vscode from 'vscode';
import { URI } from 'vscode-uri';
import { VBALSRequest } from './vba-ls-request';

export class VBAReferenceProvider implements vscode.ReferenceProvider {
    request: VBALSRequest;

    constructor(request: VBALSRequest){
        this.request = request;
    }

	async provideReferences(document: vscode.TextDocument, position: vscode.Position, context: vscode.ReferenceContext, token: vscode.CancellationToken)
	: Promise<vscode.Location[]> {
		try {
			const uri = document.uri;
			const items = await this.request.references(document, position);
			const locs = items.map(x => {
				return new vscode.Location(
					URI.file(x.filepath),
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
