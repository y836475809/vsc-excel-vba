import * as vscode from 'vscode';
import { LPSRequest } from './lsp-request';

export class VBASignatureHelpProvider implements vscode.SignatureHelpProvider {
    lpsRequest: LPSRequest;
    constructor(port: number){
        this.lpsRequest = new LPSRequest(port);
    }

    async provideSignatureHelp(document: vscode.TextDocument, position: vscode.Position, token: vscode.CancellationToken)
		: Promise<vscode.SignatureHelp | undefined>  {
		try {
            const chara = position.character;
            const data = {
                id: "SignatureHelp",
                filepaths: [document.uri.fsPath],
                line: position.line,
                chara: chara,
                text: document.getText()
            } as Hoge.RequestParam;
			const items = await this.lpsRequest.send(data) as Hoge.SignatureHelpItem[];
			if(items.length === 0){
				return undefined;
			}

			const item = items[0];
			const displaytext = item.displaytext;
			const description = item.description.replace(/\r/g, "");
			const args = item.args;
			let activeparameter = item.activeparameter;
			if(args.length === 0){
				activeparameter = 0;
				args.push({
					name: "No arguments",
					astype: ""
				});
			}

			const signatureHelp = new vscode.SignatureHelp();
			signatureHelp.activeParameter = activeparameter;
			signatureHelp.activeSignature = 0;
			const sinfo = new vscode.SignatureInformation(displaytext, description);
			sinfo.parameters = item.args.map(x => {
				return new vscode.ParameterInformation(x.name, x.astype);
			});
			signatureHelp.signatures = [
				sinfo
			];
			return signatureHelp;
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
