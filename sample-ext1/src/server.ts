import {
    createConnection,
    TextDocuments,
    Diagnostic,
    DiagnosticSeverity,
    ProposedFeatures,
    InitializeParams,
    DidChangeConfigurationNotification,
    CompletionItem,
    CompletionItemKind,
    DefinitionParams,
    TextDocumentPositionParams,
    TextDocumentSyncKind,
    InitializeResult,
    Location,
    HoverParams,
    Hover,
    MarkupContent,
    ExecuteCommandParams,
    TextDocumentChangeEvent,
    WorkspaceEdit,
    RenameParams,
  } from 'vscode-languageserver/node';
  
import { TextDocument } from 'vscode-languageserver-textdocument';
import { URI } from 'vscode-uri';
import * as fs from "fs";
import { getComdData } from "./ts-client";
import path = require('path');

const connection = createConnection(ProposedFeatures.all);
const documents: TextDocuments<TextDocument> = new TextDocuments(TextDocument);
type LpsRequest = (json: any) => Promise<string>;

export class Server {
    hasWorkspaceFolderCapability: boolean;
    changedDocSet: Set<string>;
    textDocumentMap: Map<string, TextDocument>;
    symbolKindMap:Map<string, CompletionItemKind>;
    lpsRequest: LpsRequest;

    constructor(lpsRequest: LpsRequest){
        this.hasWorkspaceFolderCapability = false;
        this.changedDocSet = new Set<string>();
        this.textDocumentMap = new Map<string, TextDocument>();
        this.symbolKindMap = new Map<string, CompletionItemKind>([
            ["Method", CompletionItemKind.Method],
            ["Field", CompletionItemKind.Field],
        ]);
        this.lpsRequest = lpsRequest;
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
            },
        };
        if (this.hasWorkspaceFolderCapability) {
            result.capabilities.workspace = {
                workspaceFolders: {
                    supported: true
                }
            };
        }
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
        const wfs = await connection.workspace.getWorkspaceFolders();
        const rootPath = (wfs && (wfs.length > 0)) ? URI.parse(wfs[0].uri).fsPath: undefined;
        if(rootPath){
            const filePaths = this.getWorkspaceFolderFiles(
                [rootPath, path.join(rootPath, ".vscode")]);
            const data = JSON.stringify({
                Id: "AddDocuments",
                FilePaths: filePaths,
                Position: 0,
                Text: ""
            }); 
            const item = JSON.parse(await this.lpsRequest(data));
            for (let index = 0; index < item.FilePaths.length; index++) {
                const uri = URI.file(item.FilePaths[index]).toString();
                const text = item.Texts[index];
                this.textDocumentMap.set(uri, TextDocument.create(
                    uri, "vb", 0, text));
            }
        }
    }

    async onRequest(method: string, params: any) {
        if(method !== "client.sendRequest"){
            return;
        }
        if(params.command === "create"){
            const uri = params.arguments?params.arguments[0]:undefined;
            const fp = URI.parse(uri).fsPath;
            const data = JSON.stringify({
                Id: "AddDocuments",
                FilePaths: [fp],
                Position: 0,
                Text: ""
            }); 
            const item = JSON.parse(await this.lpsRequest(data));
            const text = item.Texts[0];
            this.textDocumentMap.set(uri, TextDocument.create(
                uri, "vb", 0, text));
            this.changedDocSet.add(uri);
        }
        if(params.command === "delete"){
            const uri = params.arguments?params.arguments[0]:undefined;
            const fp = URI.parse(uri).fsPath;
            const data = JSON.stringify({
                Id: "DeleteDocuments",
                FilePaths: [fp],
                Position: 0,
                Text: ""
            });   
            await this.lpsRequest(data);
            this.textDocumentMap.delete(uri);
            this.changedDocSet.delete(uri);
        }
        if(params.command === "rename"){
            const renameArgs: any[] = params.arguments?params.arguments:undefined;
            if(!renameArgs){
                return;
            }
            for(const renameArg of renameArgs){
                const oldUir = renameArg.oldUir;
                const newUir = renameArg.newUir;
                const oldFsPath = URI.parse(oldUir).fsPath;
                const newFsPath = URI.parse(newUir).fsPath;
                const data = JSON.stringify({
                    Id: "RenameDocument",
                    FilePaths: [oldFsPath, newFsPath],
                    Position: 0,
                    Text: ""
                });   
                await this.lpsRequest(data);
            }
            for(const renameArg of renameArgs){
                const oldUir = renameArg.oldUir;
                const newUir = renameArg.newUir;
                const textDoc = this.textDocumentMap.get(oldUir);
                if(textDoc){
                    this.textDocumentMap.set(newUir, textDoc);
                }
                this.textDocumentMap.delete(oldUir);
                this.changedDocSet.delete(oldUir);
            }
        }
    }

    async onCompletion(_textDocumentPosition: TextDocumentPositionParams): Promise<CompletionItem[]>{
        const fp = URI.parse(_textDocumentPosition.textDocument.uri).fsPath;
        const pos = documents.get(_textDocumentPosition.textDocument.uri)?.offsetAt(_textDocumentPosition.position);
        const data = JSON.stringify({
            Id: "Completion",
            FilePaths: [fp],
            Position: pos,
            Text: documents.get(_textDocumentPosition.textDocument.uri)?.getText()
        });
        let ret = await this.lpsRequest(data);
        let res_items: any[] = JSON.parse(ret).items;
        let comlItems: CompletionItem[] = res_items.map(item => {
            const val = this.symbolKindMap.get(item.Kind);
            const kind = val?val:CompletionItemKind.Text;
            return {
                label: item.DisplayText,
                insertText: item.CompletionText,
                kind: kind
            };
        });
        return comlItems;
    }

    async onDefinition(params: DefinitionParams): Promise<Location[]> {
        const uri = params.textDocument.uri;
        const fp = URI.parse(uri).fsPath;
        const pos = documents.get(uri)?.offsetAt(params.position);
        if(!documents.get(uri)){
            return  new Array<Location>();
        }
        const data = JSON.stringify({
            Id: "Definition",
            FilePaths: [fp],
            Position: pos,
            Text: documents.get(uri)?.getText()
        });
        let ret = await this.lpsRequest(data);
        let resItems: any[] = JSON.parse(ret).items;
        const defItems: Location[] = [];
        resItems.forEach(item => {
            const defUri = URI.file(item.FilePath).toString();
            const doc = this.textDocumentMap.get(defUri);
            if(doc){        
                defItems.push(Location.create(defUri, {
                    start: doc.positionAt(item.Start),
                    end: doc.positionAt(item.End)
                }));  
            }
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
        
        const pos = doc.offsetAt(params.position);
        const data = JSON.stringify({
            Id: "Hover",
            FilePaths: [fp],
            Position: pos,
            Text: documents.get(uri)?.getText()
        });
        let ret = await this.lpsRequest(data);
        let resItems: any[] = JSON.parse(ret).items;
        if(resItems.length === 0){
            return undefined;
        }
        const item = resItems[0];
        const content: MarkupContent = {
            kind: "markdown",
            value: [
                '```vb',
                `${item.DisplayText}`,
                '```',
                '```xml',
                `${item.Description}`,
                '```',
            ].join('\n'),
        };
        return { contents: content };
    }

    async onShutdown() {
        const data = JSON.stringify({
            Id: "Shutdown",
            FilePaths: [""],
            Position: 0,
            Text: ""
        });
        await this.lpsRequest(data);
    }

    onDidChangeContent(change: TextDocumentChangeEvent<TextDocument>) {
        const uri = change.document.uri;
        this.textDocumentMap.set(uri, TextDocument.create(
            uri, "vb", 0, change.document.getText()));
        this.changedDocSet.add(change.document.uri);
    }

    async onDidSave(change: TextDocumentChangeEvent<TextDocument>) {
        const doc = change.document;
        if(this.changedDocSet.has(doc.uri)){
            const fp = URI.parse(doc.uri).fsPath;
            const data = JSON.stringify({
                Id: "ChangeDocument",
                FilePaths: [fp],
                Position: 0,
                Text: doc.getText()
            });    
            this.changedDocSet.delete(doc.uri);
            await this.lpsRequest(data);
        }
    }
}

export function startLspServer() {
    const server = new Server(getComdData);
    connection.onInitialize(server.onInitialize.bind(server));
    connection.onInitialized(server.onInitialized.bind(server));
    connection.onRequest(server.onRequest.bind(server));
    connection.onCompletion(server.onCompletion.bind(server));
    connection.onDefinition(server.onDefinition.bind(server));
    connection.onHover(server.onHover.bind(server));
    connection.onShutdown(server.onShutdown.bind(server));
    documents.onDidChangeContent(server.onDidChangeContent.bind(server));
    documents.onDidSave(server.onDidSave.bind(server));
    // Make the text document manager listen on the connection
    // for open, change and close text document events
    documents.listen(connection);
    
    // Listen on the connection
    connection.listen();
}
