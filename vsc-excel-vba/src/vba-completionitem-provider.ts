import * as vscode from 'vscode';
import { LPSRequest } from './lsp-request';

export class VBACompletionItemProvider implements vscode.CompletionItemProvider {
    lpsRequest: LPSRequest;
	symbolKindMap: Map<string, vscode.CompletionItemKind>;

    constructor(port: number){
        this.lpsRequest = new LPSRequest(port);
		this.symbolKindMap = new Map<string, vscode.CompletionItemKind>([
            ["Method",   vscode.CompletionItemKind.Method],
            ["Field",    vscode.CompletionItemKind.Field],
            ["Property", vscode.CompletionItemKind.Property],
            ["Local",    vscode.CompletionItemKind.Variable],
            ["Class",    vscode.CompletionItemKind.Class],
        ]);
    }
	
	async provideCompletionItems(document: vscode.TextDocument, position: vscode.Position, token: vscode.CancellationToken, context: vscode.CompletionContext)
	: Promise<vscode.CompletionItem[]> {
		try {
			const fp = document.uri.fsPath;
			const text = document.getText();
			if(!text){
				return [];
			}

			const data = {
				id: "Completion",
				filepaths: [fp],
				line: position.line,
				chara: position.character,
				text: text
			} as Hoge.RequestParam;
			const items = await this.lpsRequest.send(data) as Hoge.CompletionItem[];
			let comlItems = items.map(item => {
				const val = this.symbolKindMap.get(item.kind);
				const kind = val?val:vscode.CompletionItemKind.Text;
				return {
					label: item.displaytext,
					insertText: item.completiontext,
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
