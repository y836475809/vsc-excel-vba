import * as vscode from 'vscode';
import { LPSRequest } from './lsp-request';

export class VBAHoverProvider implements vscode.HoverProvider {
    lpsRequest: LPSRequest;

    constructor(port: number){
        this.lpsRequest = new LPSRequest(port);
    }

	async provideHover(document: vscode.TextDocument, position: vscode.Position, token: vscode.CancellationToken)
	: Promise<vscode.Hover | undefined> {
		try {
			const fp = document.uri.fsPath;
			const text = document.getText();
			const data = {
				id: "Hover",
				filepaths: [fp],
				line: position.line,
				chara: position.character,
				text: text
			} as Hoge.RequestParam;
			const items = await this.lpsRequest.send(data) as Hoge.CompletionItem[];
			if(items.length === 0){
				return undefined;
			}
			const item = items[0];
			const description = item.description.replace(/\r/g, "");
			const content = new vscode.MarkdownString(
				[
					'```vb',
					`${item.displaytext}`,
					'```',
					'```xml',
					`${description}`,
					'```',
				].join('\n')
			);
			return new vscode.Hover(content);		
		} catch (error) {
			if(error instanceof Error){
				vscode.window.showErrorMessage(error.message);
			}else{
				vscode.window.showErrorMessage(String(error));
			}
			return undefined;
		}
	}
}
