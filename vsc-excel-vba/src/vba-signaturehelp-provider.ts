import * as vscode from 'vscode';
import { VBALSRequest } from './vba-ls-request';

export class VBASignatureHelpProvider implements vscode.SignatureHelpProvider {
    request: VBALSRequest;

    constructor(request: VBALSRequest){
        this.request = request;
    }

    async provideSignatureHelp(document: vscode.TextDocument, position: vscode.Position, token: vscode.CancellationToken)
		: Promise<vscode.SignatureHelp | undefined>  {
		try {
			const items = await this.request.signatureHelp(document, position);
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
