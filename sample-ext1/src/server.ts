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
import { diagnosticsRequest } from './diagnostics-request';

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
            ["Property", CompletionItemKind.Property],
            ["Local", CompletionItemKind.Variable],
            ["Class", CompletionItemKind.Class],
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
        const requestMethod = method as Hoge.RequestMethod;
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
}

export function startLspServer() {
    const server = new Server();
    connection.onInitialize(server.onInitialize.bind(server));
    connection.onInitialized(server.onInitialized.bind(server));
    connection.onRequest(server.onRequest.bind(server));
    connection.onCompletion(server.onCompletion.bind(server));
    connection.onDefinition(server.onDefinition.bind(server));
    connection.onHover(server.onHover.bind(server));
    // connection.onShutdown(server.onShutdown.bind(server));
    // Make the text document manager listen on the connection
    // for open, change and close text document events
    documents.listen(connection);
    
    // Listen on the connection
    connection.listen();
}
