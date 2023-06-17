import * as vscode from 'vscode';
import { VBALSRequest } from './vba-ls-request';

export class VBACompletionItemProvider implements vscode.CompletionItemProvider {
    request: VBALSRequest;
	symbolKindMap: Map<string, vscode.CompletionItemKind>;

    constructor(request: VBALSRequest){
        this.request = request;
		this.symbolKindMap = new Map<string, vscode.CompletionItemKind>([
            ["Method",   vscode.CompletionItemKind.Method],
            ["Field",    vscode.CompletionItemKind.Field],
            ["Property", vscode.CompletionItemKind.Property],
            ["Local",    vscode.CompletionItemKind.Variable],
            ["Class",    vscode.CompletionItemKind.Class],
			["Keyword",    vscode.CompletionItemKind.Keyword],
        ]);
    }
	
	async provideCompletionItems(document: vscode.TextDocument, position: vscode.Position, token: vscode.CancellationToken, context: vscode.CompletionContext)
	: Promise<vscode.CompletionItem[]> {
		try {
			const items = await this.request.completionItems(document, position);
			let comlItems = items.map(item => {
				const val = this.symbolKindMap.get(item.kind);
				const kind = val?val:vscode.CompletionItemKind.Text;
				return {
					label: item.displaytext,
					detail: item.completiontext,
					insertText: item.displaytext,
					kind: kind
				};
			});
			return comlItems;
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
