import * as vscode from "vscode";
import { LPSRequest, MakeReqData } from './lsp-request';
import { diagnosticsRequest } from './diagnostics-request';
import { 
    VbaAttributeValidation, 
    VbaAttributeError, 
    makeAttributeDiagnostics } from "./vba-attribute-validation";

const vbaAttributeValidation = new VbaAttributeValidation();
const isVbaAttributeError = (item: any): item is Array<VbaAttributeError> => {
    if(item.length !== undefined){
        return false;
    }
    if(!(item[0] instanceof VbaAttributeError)){
        return false;
    }
    return true;
};

export class VBALSRequest {
    lpsRequest: LPSRequest;
    diagnosticCollection: vscode.DiagnosticCollection;

    constructor(port: number){
        this.lpsRequest = new LPSRequest(port);
        this.diagnosticCollection = vscode.languages.createDiagnosticCollection("vba");
    }

    async addDocuments(uris: vscode.Uri[]) {
        const data = MakeReqData.addDocuments(uris);
        await this.lpsRequest.send(data);
    }
    async deleteDocuments(uris: vscode.Uri[]) {  
        const data = MakeReqData.deleteDocuments(uris);
        await this.lpsRequest.send(data);
    }
    async renameDocument(params: Hoge.RequestRenameParam[]) {
        if(!params){
            return;
        }
        const dataList = MakeReqData.renameDocuments(params);
        for(const data of dataList){ 
            await this.lpsRequest.send(data);
        }
    }
    async changeDocument(document: vscode.TextDocument) {
        const uri = document.uri;
        if(!uri){
            return;
        }
        const text = document.getText();
        if(!text){
            return;
        }
        const fsPath = uri.fsPath;
        const data = MakeReqData.changeDocument(uri, text);
        await this.lpsRequest.send(data);

        const items = await diagnosticsRequest(fsPath, this.lpsRequest);
        
        try {
            vbaAttributeValidation.validate(uri, text);
        } catch (error) {
            if(isVbaAttributeError(error)){
                const attrDiagnostics = makeAttributeDiagnostics(error);
                items.push(...attrDiagnostics);
            }
        }

        this.diagnosticCollection.set(document.uri, items);
    }
    async diagnostic(document: vscode.TextDocument) {
        const uri = document.uri;
        if(!uri){
            return;
        }
        const text = document.getText();
        if(!text){
            return;
        }
        const fsPath = uri.fsPath;
        const items = await diagnosticsRequest(fsPath, this.lpsRequest);
        
        try {
            vbaAttributeValidation.validate(uri, text);
        } catch (error) {
            if(isVbaAttributeError(error)){
                const attrDiagnostics = makeAttributeDiagnostics(error);
                items.push(...attrDiagnostics);
            }
        } 

        this.diagnosticCollection.set(document.uri, items);
    }
    async reset() {
        const data = MakeReqData.reset();
        await this.lpsRequest.send(data);
    }
    async isReady() {
        const data = MakeReqData.isReady();
        await this.lpsRequest.send(data);
    }
    async shutdown() {
        const data = MakeReqData.shutdown();
        await this.lpsRequest.send(data);
    }
}