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
import { EphemeralKeyInfo } from 'tls';

const connection = createConnection(ProposedFeatures.all);
const documents: TextDocuments<TextDocument> = new TextDocuments(TextDocument);
type LpsRequest = (json: any) => Promise<any>;

export class Server {
    hasWorkspaceFolderCapability: boolean;
    symbolKindMap:Map<string, CompletionItemKind>;
    lpsRequest: LpsRequest;

    constructor(lpsRequest: LpsRequest){
        this.hasWorkspaceFolderCapability = false;
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
            const data = {
                Id: "AddDocuments",
                FilePaths: filePaths,
                Position: 0,
                Text: ""
            }; 
            await this.lpsRequest(data);
        }
    }

    async onRequest(method: string, params: any) {
        if(method !== "client.sendRequest"){
            return;
        }
        if(params.command === "create"){
            const uris: string[] = params.arguments?params.arguments:undefined;
            const filePaths = uris.map(uri => URI.parse(uri).fsPath);
            const data = {
                Id: "AddDocuments",
                FilePaths: filePaths,
                Position: 0,
                Text: ""
            }; 
            await this.lpsRequest(data);
        }
        if(params.command === "delete"){
            const uris:string[] = params.arguments?params.arguments:undefined;
            const fsPaths = uris.map(uri => {
                return URI.parse(uri).fsPath;
            });
            const data = {
                Id: "DeleteDocuments",
                FilePaths: fsPaths,
                Position: 0,
                Text: ""
            };   
            await this.lpsRequest(data);
        }
        if(params.command === "rename"){
            const renameArgs: any[] = params.arguments?params.arguments:undefined;
            if(!renameArgs){
                return;
            }
            for(const renameArg of renameArgs){
                const oldUri = renameArg.oldUri;
                const newUri = renameArg.newUri;
                const oldFsPath = URI.parse(oldUri).fsPath;
                const newFsPath = URI.parse(newUri).fsPath;
                const data = {
                    Id: "RenameDocument",
                    FilePaths: [oldFsPath, newFsPath],
                    Position: 0,
                    Text: ""
                };   
                await this.lpsRequest(data);
            }
        }
        if(params.command === "changeDocument"){
            const uri = params.arguments?params.arguments[0]:undefined;
            if(!uri){
                return;
            }
            const fsPath = URI.parse(uri).fsPath;
            const data = {
                Id: "ChangeDocument",
                FilePaths: [fsPath],
                Position: 0,
                Text: documents.get(uri)?.getText()
            };
            await this.lpsRequest(data);
        }
    }

    async onCompletion(_textDocumentPosition: TextDocumentPositionParams): Promise<CompletionItem[]>{
        const fp = URI.parse(_textDocumentPosition.textDocument.uri).fsPath;
        const pos = documents.get(_textDocumentPosition.textDocument.uri)?.offsetAt(_textDocumentPosition.position);
        const data = {
            Id: "Completion",
            FilePaths: [fp],
            Position: pos,
            Text: documents.get(_textDocumentPosition.textDocument.uri)?.getText()
        };
        let ret:any = await this.lpsRequest(data);
        let res_items: any[] = ret.items;
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
        const data = {
            Id: "Definition",
            FilePaths: [fp],
            Position: pos,
            Text: documents.get(uri)?.getText()
        };
        let ret:any = await this.lpsRequest(data);
        let resItems: any[] = ret.items;
        const defItems: Location[] = [];
        resItems.forEach(item => {
            const defUri = URI.file(item.FilePath).toString();
            const start = item.Start;
            const end = item.End;
            defItems.push(Location.create(defUri, {
                start: {line: start.Line, character: start.Character},
                end: {line: end.Line, character: end.Character}
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
        
        const pos = doc.offsetAt(params.position);
        const data = {
            Id: "Hover",
            FilePaths: [fp],
            Position: pos,
            Text: documents.get(uri)?.getText()
        };
        let ret:any = await this.lpsRequest(data);
        let resItems: any[] = ret.items;
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
        const data = {
            Id: "Shutdown",
            FilePaths: [""],
            Position: 0,
            Text: ""
        };
        await this.lpsRequest(data);
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
    // Make the text document manager listen on the connection
    // for open, change and close text document events
    documents.listen(connection);
    
    // Listen on the connection
    connection.listen();
}
