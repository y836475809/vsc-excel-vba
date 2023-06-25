import * as vscode from 'vscode';
import { URI } from 'vscode-uri';
import { VBALSRequest } from './vba-ls-request';

export class VBADefinitionProvider implements vscode.DefinitionProvider {
    request: VBALSRequest;

    constructor(request: VBALSRequest){
        this.request = request;
    }

	async provideDefinition(document: vscode.TextDocument, position: vscode.Position, token: vscode.CancellationToken)
	: Promise<vscode.Location[] | undefined> {
		try {	
			const items = await this.request.definition(document.uri, position) as VEV.DefinitionItem[];
			const locations: vscode.Location[] = [];
			items.forEach(item => {
				const uri = URI.file(item.filepath);
				const start = item.start;
				const end = item.end;
				locations.push(new vscode.Location(uri, new vscode.Range(
					start.line, start.character,
					end.line, end.character
				)));
			});
			return locations;
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
