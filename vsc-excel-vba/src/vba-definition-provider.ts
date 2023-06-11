import * as vscode from 'vscode';
import { URI } from 'vscode-uri';
import { LPSRequest } from './lsp-request';

export class VBADefinitionProvider implements vscode.DefinitionProvider {
    lpsRequest: LPSRequest;
    constructor(port: number){
        this.lpsRequest = new LPSRequest(port);
    }

	async provideDefinition(document: vscode.TextDocument, position: vscode.Position, token: vscode.CancellationToken)
	: Promise<vscode.Location[] | undefined> {
		try {
			const fp = document.uri.fsPath;
			const line = position.line;
			const chara = position.character;

			const data = {
				id: "Definition",
				filepaths: [fp],
				line: line,
				chara: chara,
				text: ""
			} as Hoge.RequestParam;
			
			const items = await this.lpsRequest.send(data) as Hoge.DefinitionItem[];
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
