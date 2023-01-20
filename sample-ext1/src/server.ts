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
import { LPSRequest } from "./lsp-request";
import path = require('path');
import { EphemeralKeyInfo } from 'tls';

const connection = createConnection(ProposedFeatures.all);
const documents: TextDocuments<TextDocument> = new TextDocuments(TextDocument);

async function getSetting(){
    const settings = await connection.workspace.getConfiguration(
        [{ section: "sample-ext1" }]) as any[];
    return settings[0];
}

export class Server {
    hasWorkspaceFolderCapability: boolean;
    symbolKindMap:Map<string, CompletionItemKind>;
    lpsRequest!: LPSRequest;

    constructor(){
        this.hasWorkspaceFolderCapability = false;
        this.symbolKindMap = new Map<string, CompletionItemKind>([
            ["Method", CompletionItemKind.Method],
            ["Field", CompletionItemKind.Field],
        ]);
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

        const setting = await getSetting();
        const port = setting.serverPort as number;
        this.lpsRequest = new LPSRequest(port);
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
            } as Hoge.Command;
            await this.lpsRequest.send(data);
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
            } as Hoge.Command;  
            await this.lpsRequest.send(data);
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
                } as Hoge.Command;  
                await this.lpsRequest.send(data);
            }
        }
        if(params.command === "changeDocument"){
            const uri = params.arguments?params.arguments[0]:undefined;
            if(!uri){
                return;
            }
            const doc = documents.get(uri);
            if(!doc){
                return;
            }
            const fsPath = URI.parse(uri).fsPath;
            const data = {
                Id: "ChangeDocument",
                FilePaths: [fsPath],
                Position: 0,
                Text: doc.getText()
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
        const pos = doc.offsetAt(_textDocumentPosition.position);
        const text = doc.getText();
        const data = {
            Id: "Completion",
            FilePaths: [fp],
            Position: pos,
            Text: text
        } as Hoge.Command;
        let ret = await this.lpsRequest.send(data) as Hoge.CompletionItems;
        let res_items = ret.items;
        let comlItems = res_items.map(item => {
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
        } as Hoge.Command;
        let ret = await this.lpsRequest.send(data) as Hoge.DefinitionItems;
        let resItems = ret.items;
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
        } as Hoge.Command;
        let ret = await this.lpsRequest.send(data) as Hoge.CompletionItems;
        let resItems = ret.items;
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
            FilePaths: [],
            Position: 0,
            Text: ""
        } as Hoge.Command;
        await this.lpsRequest.send(data);
    }
}

export function startLspServer() {
    const server = new Server();
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
