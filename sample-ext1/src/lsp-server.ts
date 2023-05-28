import {
    createConnection,
    TextDocuments,
    ProposedFeatures,
    InitializeParams,
    CompletionItem,
    CompletionItemKind,
    DefinitionParams,
    TextDocumentPositionParams,
    TextDocumentSyncKind,
    InitializeResult,
    Location,
    HoverParams,
    Hover,
    ReferenceParams,
    MarkupContent,
  } from 'vscode-languageserver/node';
  
import { TextDocument } from 'vscode-languageserver-textdocument';
import { URI } from 'vscode-uri';
import * as fs from "fs";
import { LPSRequest } from "./lsp-request";
import path = require('path');
import { diagnosticsRequest } from './diagnostics-request';
import { Logger } from "./logger";
import { 
    VbaAttributeValidation, 
    VbaAttributeError, 
    makeAttributeDiagnostics } from "./vba-attribute-validation";

const connection = createConnection(ProposedFeatures.all);
const documents: TextDocuments<TextDocument> = new TextDocuments(TextDocument);
let logger: Logger;
const vbaAttributeValidation = new VbaAttributeValidation();

const defaultPort = 9088;

export class LSPServer {
    hasWorkspaceFolderCapability: boolean;
    symbolKindMap:Map<string, CompletionItemKind>;
    lpsRequest!: LPSRequest;
    port: number;

    constructor(){
        this.hasWorkspaceFolderCapability = false;
        this.symbolKindMap = new Map<string, CompletionItemKind>([
            ["Method", CompletionItemKind.Method],
            ["Field", CompletionItemKind.Field],
            ["Property", CompletionItemKind.Property],
            ["Local", CompletionItemKind.Variable],
            ["Class", CompletionItemKind.Class],
        ]);
        this.port = defaultPort;
        
        logger = new Logger((msg: string) => {
            connection.console.log(msg);
        });
    }

    async initLSPRequest() {
        if(!this.lpsRequest){
            this.lpsRequest = new LPSRequest(this.port);
        } 
    }

    private getPort(params: InitializeParams): number{
        const args = params.initializationOptions?.arguments as string[];
        if(!args){
            logger.error(`initializationOptions.arguments is non, use default port: ${defaultPort}`);
            return defaultPort;
        }
        const port = Number(args[0]);
        if(isNaN(port)){
            logger.error(`initializationOptions.arguments[0] is NaN, use default port: ${defaultPort}`);
            return defaultPort;
        }
        return port;
    }

    onInitialize(params: InitializeParams){
        let capabilities = params.capabilities;
        this.hasWorkspaceFolderCapability = !!(
            capabilities.workspace && !!capabilities.workspace.workspaceFolders
        );
        const result: InitializeResult = {
            capabilities: {
                textDocumentSync: TextDocumentSyncKind.Incremental,
                completionProvider: {
                    resolveProvider: false,
                },
                definitionProvider: true,
                hoverProvider: true,
                referencesProvider: true,
            },
        };
        if (this.hasWorkspaceFolderCapability) {
            result.capabilities.workspace = {
                workspaceFolders: {
                    supported: true
                }
            };
        }
        this.port = this.getPort(params);
        return result;
    }

    getWorkspaceFolderFiles(wsPaths: string[]): string[] {
        let filePaths: string[] = [];
        for(const wsPath of wsPaths){
            const fsPaths = fs.readdirSync(wsPath, { withFileTypes: true })
            .filter(dirent => {
                return dirent.isFile() 
                    && (dirent.name.endsWith('.bas') || dirent.name.endsWith('.cls'));
            }).map(dirent => path.join(wsPath, dirent.name));
            filePaths = filePaths.concat(fsPaths);
        }
        return filePaths;
    }

    async onInitialized() {
        if (this.hasWorkspaceFolderCapability) {
            connection.workspace.onDidChangeWorkspaceFolders(_event => {
                connection.console.log('Workspace folder change event received.');
            });
        }

        await this.initLSPRequest();
    }

    async onRequest(method: string, params: any) {
        await this.initLSPRequest();

        const requestMethod = method as Hoge.RequestMethod;
        logger.info(`onRequest: ${requestMethod}`);
        if(requestMethod === "createFiles"){
            const uris = params.uris as string[];
            const filePaths = uris.map(uri => URI.parse(uri).fsPath);
            const data = {
                id: "AddDocuments",
                filepaths: filePaths,
                line: 0,
                chara: 0,
                text: ""
            } as Hoge.Command;
            await this.lpsRequest.send(data);
        }
        if(requestMethod === "deleteFiles"){
            const uris = params.uris as string[];
            const fsPaths = uris.map(uri => {
                return URI.parse(uri).fsPath;
            });
            const data = {
                id: "DeleteDocuments",
                filepaths: fsPaths,
                line: 0,
                chara: 0,
                text: ""
            } as Hoge.Command;  
            await this.lpsRequest.send(data);
        }
        if(requestMethod === "renameFiles"){
            const renameArgs = params.renameParams as Hoge.RequestRenameParam[];
            if(!renameArgs){
                return;
            }
            for(const renameArg of renameArgs){
                const oldUri = renameArg.olduri;
                const newUri = renameArg.newuri;
                const oldFsPath = URI.parse(oldUri).fsPath;
                const newFsPath = URI.parse(newUri).fsPath;
                const data = {
                    id: "RenameDocument",
                    filepaths: [oldFsPath, newFsPath],
                    line: 0,
                    chara: 0,
                    text: ""
                } as Hoge.Command;  
                await this.lpsRequest.send(data);
            }
        }
        if(requestMethod === "changeText"){
            const uri = params.uri as string;
            if(!uri){
                return;
            }
            const doc = documents.get(uri);
            if(!doc){
                return;
            }
            const fsPath = URI.parse(uri).fsPath;
            const data = {
                id: "ChangeDocument",
                filepaths: [fsPath],
                position: 0,
                line: 0,
                chara: 0,
                text: doc.getText()
            } as Hoge.Command;
            await this.lpsRequest.send(data);

            const items = await diagnosticsRequest(doc, fsPath, this.lpsRequest);
            
            try {
                vbaAttributeValidation.validate(URI.parse(uri), doc.getText());
            } catch (error) {
                if(error instanceof Array<VbaAttributeError>){
                    const attrDiagnostics = makeAttributeDiagnostics(error);
                    items.push(...attrDiagnostics);
                }
            }
            connection.sendDiagnostics({uri: doc.uri, diagnostics: items});
        }
        if(requestMethod === "diagnostics"){
            const uri = params.uri as string;
            if(!uri){
                return;
            }
            const doc = documents.get(uri);
            if(!doc){
                return;
            }
            const fsPath = URI.parse(uri).fsPath;
            const items = await diagnosticsRequest(doc, fsPath, this.lpsRequest);
            
            try {
                vbaAttributeValidation.validate(URI.parse(uri), doc.getText());
            } catch (error) {
                if(error instanceof Array<VbaAttributeError>){
                    const attrDiagnostics = makeAttributeDiagnostics(error);
                    items.push(...attrDiagnostics);
                }
            } 

            connection.sendDiagnostics({uri: doc.uri, diagnostics: items});
        }
        if(requestMethod === "reset"){
            const data = {
                id: "Reset",
                filepaths: [],
                line: 0,
                chara: 0,
                text: ""
            } as Hoge.Command;  
            await this.lpsRequest.send(data);
        }
        if(requestMethod === "IsReady"){
            const data = {
                id: "IsReady",
                filepaths: [],
                line: 0,
                chara: 0,
                text: ""
            } as Hoge.Command;
            await this.lpsRequest.send(data);
        }
        if(requestMethod === "Shutdown"){
            const data = {
                id: "Shutdown",
                filepaths: [],
                line: 0,
                chara: 0,
                text: ""
            } as Hoge.Command;
            await this.lpsRequest.send(data);
        }
    }

    async onCompletion(_textDocumentPosition: TextDocumentPositionParams): Promise<CompletionItem[]>{
        const fp = URI.parse(_textDocumentPosition.textDocument.uri).fsPath;
        const doc = documents.get(_textDocumentPosition.textDocument.uri);
        if(!doc){
            return [];
        }
        const text = doc.getText();
        const data = {
            id: "Completion",
            filepaths: [fp],
            line: _textDocumentPosition.position.line,
            chara: _textDocumentPosition.position.character,
            text: text
        } as Hoge.Command;
        const items = await this.lpsRequest.send(data) as Hoge.CompletionItem[];
        let comlItems = items.map(item => {
            const val = this.symbolKindMap.get(item.kind);
            const kind = val?val:CompletionItemKind.Text;
            return {
                label: item.displaytext,
                insertText: item.completiontext,
                kind: kind
            };
        });
        return comlItems;
    }

    async onDefinition(params: DefinitionParams): Promise<Location[]> {
        const uri = params.textDocument.uri;
        const fp = URI.parse(uri).fsPath;
        if(!documents.get(uri)){
            return  new Array<Location>();
        }
        const data = {
            id: "Definition",
            filepaths: [fp],
            line: params.position.line,
            chara: params.position.character,
            text: documents.get(uri)?.getText()
        } as Hoge.Command;
        const items = await this.lpsRequest.send(data) as Hoge.DefinitionItem[];
        const defItems: Location[] = [];
        items.forEach(item => {
            const defUri = URI.file(item.filepath).toString();
            const start = item.start;
            const end = item.end;
            defItems.push(Location.create(defUri, {
                start: {line: start.line, character: start.character},
                end: {line: end.line, character: end.character}
            }));
        });
        return defItems;
    }

    async onHover(params: HoverParams): Promise<Hover | undefined> {
        const uri = params.textDocument.uri;
        const fp = URI.parse(uri).fsPath;
        const doc = documents.get(uri);
        if(!doc){
            return undefined;
        }
        const data = {
            id: "Hover",
            filepaths: [fp],
            line: params.position.line,
            chara: params.position.character,
            text: documents.get(uri)?.getText()
        } as Hoge.Command;
        const items = await this.lpsRequest.send(data) as Hoge.CompletionItem[];
        if(items.length === 0){
            return undefined;
        }
        const item = items[0];
        const description = item.description.replace(/\r/g, "");
        const content: MarkupContent = {
            kind: "markdown",
            value: [
                '```vb',
                `${item.displaytext}`,
                '```',
                '```xml',
                `${description}`,
                '```',
            ].join('\n'),
        };
        return { contents: content };
    }

    async onShutdown() {
        const data = {
            id: "Shutdown",
            filepaths: [],
            line: 0,
            chara: 0,
            text: ""
        } as Hoge.Command;
        await this.lpsRequest.send(data);
    }

    async onReferences(params: ReferenceParams): Promise<Location[]> {
        const uri = params.textDocument.uri;
        const fp = URI.parse(uri).fsPath;
        const position = params.position;
        const data = {
            id: "References",
            filepaths: [fp],
            line: position.line,
            chara: position.character,
            text: ""
        } as Hoge.Command;
        const items = await this.lpsRequest.send(data) as Hoge.ReferencesItem[];
        const locs = items.map(x => {
            return {
                uri: URI.file(x.filepath).toString(),
                range: {
                    start:{ line: x.start.line, character: x.start.character },
                    end: { line: x.end.line, character: x.end.character }
                }    
            };
        });
        return locs;
    }
}

export function startLSPServer() {
    const server = new LSPServer();
    connection.onInitialize(server.onInitialize.bind(server));
    connection.onInitialized(server.onInitialized.bind(server));
    connection.onRequest(server.onRequest.bind(server));
    connection.onCompletion(server.onCompletion.bind(server));
    connection.onDefinition(server.onDefinition.bind(server));
    connection.onHover(server.onHover.bind(server));
    connection.onReferences(server.onReferences.bind(server));
    // connection.onShutdown(server.onShutdown.bind(server));
    // Make the text document manager listen on the connection
    // for open, change and close text document events
    documents.listen(connection);
    
    // Listen on the connection
    connection.listen();
}
