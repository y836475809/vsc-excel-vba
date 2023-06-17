import * as vscode from 'vscode';
import { VBALSRequest } from './vba-ls-request';

export class VBAHoverProvider implements vscode.HoverProvider {
    request: VBALSRequest;

    constructor(request: VBALSRequest){
        this.request = request;
    }

	async provideHover(document: vscode.TextDocument, position: vscode.Position, token: vscode.CancellationToken)
	: Promise<vscode.Hover | undefined> {
		try {
			const items = await this.request.hover(document, position);
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
