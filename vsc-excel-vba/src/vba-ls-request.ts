import * as vscode from "vscode";
import * as path from "path";
import { LPSRequest, MakeReqData, RenameParam } from './lsp-request';
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
    srcDir: string;
    diagnosticCollection: vscode.DiagnosticCollection;

    constructor(port: number){
        this.srcDir = "";
        this.lpsRequest = new LPSRequest(port);
        this.diagnosticCollection = vscode.languages.createDiagnosticCollection("vba");
    }

    private isSrcFile(uri: vscode.Uri): boolean{
        const ext = path.extname(uri.fsPath);
        return uri.scheme === "file" 
            && path.dirname(uri.fsPath).startsWith(this.srcDir)
            && (ext === ".bas" || ext === ".cls");
    }

    private getSrcUris(uris: vscode.Uri[]): vscode.Uri[] {
        if(this.srcDir === ""){
            throw new Error(`srcDir is Empyt`);
        }
        const srcUris = 
            uris.filter(uri => this.isSrcFile(uri));
        return srcUris;
    }

    async addDocuments(uris: vscode.Uri[]) {
        const srcUris = this.getSrcUris(uris);
        if(!srcUris.length){
            return;
        }
        const data = MakeReqData.addDocuments(srcUris);
        await this.lpsRequest.send(data);
    }
    async addVBADefines(uris: vscode.Uri[]) {
        const vbadUtis = uris.filter(uri => uri.fsPath.endsWith(".d.vb"));
        if(!vbadUtis.length){
            return;
        }
        const data = MakeReqData.addDocuments(vbadUtis);
        await this.lpsRequest.send(data);
    }
    async deleteDocuments(uris: vscode.Uri[]) {  
        const srcUris = this.getSrcUris(uris);
        if(!srcUris.length){
            return;
        }
        srcUris.forEach(x => {
            this.diagnosticCollection.delete(x);
        });

        const data = MakeReqData.deleteDocuments(srcUris);
        await this.lpsRequest.send(data);
    }
    async renameDocument(params: RenameParam[]) {
        if(!params){
            return;
        }

        const files = params
            .filter(x => this.isSrcFile(x.oldUri) && this.isSrcFile(x.newUri))
            .map(x => {
                return {
                    olduri: x.oldUri.toString(),
                    newuri: x.newUri.toString()
                };
            });
        if(!files.length){
            const movedFiles = params
                .filter(x => this.isSrcFile(x.oldUri))
                .filter(x => !this.isSrcFile(x.newUri))
                .map(x => x.oldUri);
            await this.deleteDocuments(movedFiles);
            return;
        }

        params.forEach(x => {   
            this.diagnosticCollection.delete(x.oldUri);
        });

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
        const srcUris = this.getSrcUris([uri]);
        if(!srcUris.length){
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
        const srcUris = this.getSrcUris([uri]);
        if(!srcUris.length){
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

    async definition(uri: vscode.Uri, position: vscode.Position)
        : Promise<Hoge.DefinitionItem[]> {
        const srcUris = this.getSrcUris([uri]);
        if(!srcUris.length){
            return [];
        }

        const data = {
            id: "Definition",
            filepaths: [uri.fsPath],
            line: position.line,
            chara: position.character,
            text: ""
        } as Hoge.RequestParam;   
        return await this.lpsRequest.send(data) as Hoge.DefinitionItem[];
    }
    
    async hover(document: vscode.TextDocument, position: vscode.Position)
        : Promise<Hoge.CompletionItem[]> {
        const srcUris = this.getSrcUris([document.uri]);
        if(!srcUris.length){
            return [];
        }

        const data = {
            id: "Hover",
            filepaths: [document.uri.fsPath],
            line: position.line,
            chara: position.character,
            text: document.getText()
        } as Hoge.RequestParam;   
        return await this.lpsRequest.send(data) as Hoge.CompletionItem[];
    }

    async completionItems(document: vscode.TextDocument, position: vscode.Position)
        : Promise<Hoge.CompletionItem[]> {
        const srcUris = this.getSrcUris([document.uri]);
        if(!srcUris.length){
            return [];
        }

        const data = {
            id: "Completion",
            filepaths: [document.uri.fsPath],
            line: position.line,
            chara: position.character,
            text: document.getText()
        } as Hoge.RequestParam;  
        return await this.lpsRequest.send(data) as Hoge.CompletionItem[];
    }

    async references(document: vscode.TextDocument, position: vscode.Position)
        : Promise<Hoge.ReferencesItem[]> {
        const srcUris = this.getSrcUris([document.uri]);
        if(!srcUris.length){
            return [];
        }

        const data = {
            id: "References",
            filepaths: [document.uri.fsPath],
            line: position.line,
            chara: position.character,
            text: ""
        } as Hoge.RequestParam; 
        return await this.lpsRequest.send(data) as Hoge.ReferencesItem[];
    }

    async signatureHelp(document: vscode.TextDocument, position: vscode.Position)
        : Promise<Hoge.SignatureHelpItem[]> {
        const srcUris = this.getSrcUris([document.uri]);
        if(!srcUris.length){
            return [];
        }

        const data = {
            id: "SignatureHelp",
            filepaths: [document.uri.fsPath],
            line: position.line,
            chara: position.character,
            text: document.getText()
        } as Hoge.RequestParam;
        return await this.lpsRequest.send(data) as Hoge.SignatureHelpItem[];
    }
}