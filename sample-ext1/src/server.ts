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
    MarkupContent
  } from 'vscode-languageserver/node';
  
import { TextDocument } from 'vscode-languageserver-textdocument';
import { URI } from 'vscode-uri';
import * as fs from "fs";
import { getComdData } from "./ts-client";
import path = require('path');

const connection = createConnection(ProposedFeatures.all);
const documents: TextDocuments<TextDocument> = new TextDocuments(TextDocument);
let hasWorkspaceFolderCapability: boolean = false;
const changedDocSet = new Set<string>();
const textDocumentMap = new Map<string, TextDocument>();

connection.onInitialize((params: InitializeParams) => {
    let capabilities = params.capabilities;
    hasWorkspaceFolderCapability = !!(
        capabilities.workspace && !!capabilities.workspace.workspaceFolders
    );

    const result: InitializeResult = {
        capabilities: {
            textDocumentSync: TextDocumentSyncKind.Incremental,
            // // Tell the client that this server supports code completion.
            completionProvider: {
                resolveProvider: false,
                // triggerCharacters: ["."]
                // triggerCharacters: ['--'],
                // completionItem: 
            },
            definitionProvider: true,
            // // Tell the client that this server supports hover.
            hoverProvider: true,
        },
    };
    if (hasWorkspaceFolderCapability) {
        result.capabilities.workspace = {
            workspaceFolders: {
                supported: true
            }
        };
    }
    return result;
});

// let all_bk:TextDocument[] = []
connection.onInitialized(async () => {
    // if (hasConfigurationCapability) {
    //     // Register for all configuration changes.
    //     connection.client.register(DidChangeConfigurationNotification.type, undefined);
    // }
    if (hasWorkspaceFolderCapability) {
        connection.workspace.onDidChangeWorkspaceFolders(_event => {
            connection.console.log('Workspace folder change event received.');
        });
        connection.workspace.onDidDeleteFiles(params => {
            params.files.forEach(file => {
                textDocumentMap.delete(file.uri);
            });
        });
    }
    const wfs = await connection.workspace.getWorkspaceFolders();
    const rootPath = (wfs && (wfs.length > 0)) ? URI.parse(wfs[0].uri).fsPath: undefined;
    if(rootPath){
        const fsPaths = fs.readdirSync(rootPath, { withFileTypes: true })
            .filter(dirent => {
                return dirent.isFile() 
                    && (dirent.name.endsWith('.bas') || dirent.name.endsWith('.cls'));
            }).map(dirent => path.join(rootPath, dirent.name));
        // const uris = await vscode.workspace.findFiles("*.{bas,cls}");
        // const fsPaths = uris.map(uri => uri.fsPath);
        for (let index = 0; index < fsPaths.length; index++) {
            const fp = fsPaths[index];
            const data = JSON.stringify({
                Id: "AddDocument",
                FilePath: fp,
                Position: 0,
                Text: ""
            });
            const uri = URI.file(fp).toString();
            // const buf = fs.readFileSync(fp);
            // const content = iconv.decode(buf, "Shift_JIS");
            const item = JSON.parse(await getComdData(data));
            textDocumentMap.set(uri, TextDocument.create(
                uri, "vb", 0, item.Text));
        }
    }
});

documents.onDidChangeContent(change => {
    let m = 0;
    // let all_bk = documents.all();
    console.log("change=", change);
    // change.document.getText();
    // alidateTextDocument(change.document);
    // change.document.positionAt()
    const uri = change.document.uri;
    textDocumentMap.set(uri, TextDocument.create(
        uri, "vb", 0, change.document.getText()));
    changedDocSet.add(change.document.uri);
});

documents.onDidSave(async change => {
    console.log("onDidSave change=", change);
    const doc = change.document;
    if(changedDocSet.has(doc.uri)){
        const fp = URI.parse(doc.uri).fsPath;
        const data = JSON.stringify({
            Id: "ChangeDocument",
            FilePath: fp,
            Position: 0,
            Text: doc.getText()
        });    
        changedDocSet.delete(doc.uri);
        await getComdData(data);
    }
});


connection.onCompletion(async (_textDocumentPosition: TextDocumentPositionParams): Promise<CompletionItem[]> => {
    // The pass parameter contains the position of the text document in
    // which code complete got requested. For the example we ignore this
    // info and always provide the same completion items.
        // const json_data = JSON.stringify({
        //     cmd: "OK",
        //     line: 10,
        //     col:25,
        //     uri: _textDocumentPosition.textDocument.uri
        // });
        // const data = JSON.stringify({
        //     id: "text",
        //     json_string: json_data
        // });
       
    const fp = URI.parse(_textDocumentPosition.textDocument.uri).fsPath;
    const pos = documents.get(_textDocumentPosition.textDocument.uri)?.offsetAt(_textDocumentPosition.position);
    const data = JSON.stringify({
        Id: "Completion",
        FilePath: fp,
        Position: pos,
        Text: documents.get(_textDocumentPosition.textDocument.uri)?.getText()
    });
    let ret = await getComdData(data);
    let res_items: any[] = JSON.parse(ret).items;
    let comlItems: CompletionItem[] = res_items.map(item => {
        return {
            label: item.DisplayText,
            insertText: item.CompletionText,
            kind: CompletionItemKind.Text
        };
    });
    return comlItems;
    // return [
    // {
    //     label: 'TypeScript',
    //     kind: CompletionItemKind.Text,
    //     data: 1
    // },
    // {
    //     label: 'JavaScript',
    //     kind: CompletionItemKind.Text,
    //     data: 2
    // }
    // ];
});
connection.onCompletionResolve((item: CompletionItem): CompletionItem => {
    // if (item.data === 1) {
    //     item.detail = 'TypeScript details';
    //     item.documentation = 'TypeScript documentation';
    // } else if (item.data === 2) {
    //     item.detail = 'JavaScript details';
    //     item.documentation = 'JavaScript documentation';
    // }
    return item;
});
// connection.onDeclaration
connection.onDefinition(async (params: DefinitionParams): Promise<Location[]> => {
    console.log(params);

    const uri = params.textDocument.uri;
    const fp = URI.parse(uri).fsPath;
    const pos = documents.get(uri)?.offsetAt(params.position);
    if(!documents.get(uri)){
        return  new Array<Location>();
    }
    const data = JSON.stringify({
        Id: "Definition",
        FilePath: fp,
        Position: pos,
        Text: documents.get(uri)?.getText()
    });
    let ret = await getComdData(data);
    let resItems: any[] = JSON.parse(ret).items;
    const defItems: Location[] = [];
    resItems.forEach(item => {
        const defUri = URI.file(item.FilePath).toString();
        const doc = textDocumentMap.get(defUri);
        if(doc){        
            defItems.push(Location.create(defUri, {
                start: doc.positionAt(item.Start),
                end: doc.positionAt(item.End)
            }));  
        }
    });
    return defItems;
});

connection.onHover(async (params: HoverParams): Promise<Hover | undefined> => {
    const uri = params.textDocument.uri;
    const fp = URI.parse(uri).fsPath;
    const doc = documents.get(uri);
    if(!doc){
        return undefined;
    }
    
    const pos = doc.offsetAt(params.position);
    const data = JSON.stringify({
        Id: "Hover",
        FilePath: fp,
        Position: pos,
        Text: documents.get(uri)?.getText()
    });
    let ret = await getComdData(data);
    let resItems: any[] = JSON.parse(ret).items;
    // const contents: string[] = resItems.map(item => {
    //     return item.DisplayText;
    // });
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
    // return { contents: contents };
});
// connection.onExit(async ()=>{
//     const data = JSON.stringify({
//         Id: "Exit",
//         FilePath: "",
//         Position: 0,
//         Text: ""
//     });
//     await getComdData(data);
// });
connection.onShutdown(async ()=>{
    const data = JSON.stringify({
        Id: "Shutdown",
        FilePath: "",
        Position: 0,
        Text: ""
    });
    await getComdData(data);
});

// Make the text document manager listen on the connection
// for open, change and close text document events
documents.listen(connection);

// Listen on the connection
connection.listen();