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
    TextDocumentPositionParams,
    TextDocumentSyncKind,
    InitializeResult
  } from 'vscode-languageserver/node';
  
import { TextDocument } from 'vscode-languageserver-textdocument';

import { getComdData } from "./ts-client";

const connection = createConnection(ProposedFeatures.all);
const documents: TextDocuments<TextDocument> = new TextDocuments(TextDocument);
let hasWorkspaceFolderCapability: boolean = false;

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
            // // Tell the client that this server supports hover.
            // hoverProvider: true,
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

connection.onInitialized(() => {
    // if (hasConfigurationCapability) {
    //     // Register for all configuration changes.
    //     connection.client.register(DidChangeConfigurationNotification.type, undefined);
    // }
    if (hasWorkspaceFolderCapability) {
        connection.workspace.onDidChangeWorkspaceFolders(_event => {
            connection.console.log('Workspace folder change event received.');
        });
    }
});

documents.onDidChangeContent(change => {
    let m = 0;
    // alidateTextDocument(change.document);
});

connection.onCompletion((_textDocumentPosition: TextDocumentPositionParams): CompletionItem[] => {
    // The pass parameter contains the position of the text document in
    // which code complete got requested. For the example we ignore this
    // info and always provide the same completion items.
        const json_data = JSON.stringify({
            cmd: "OK",
            line: 10,
            col:25,
            uri: _textDocumentPosition.textDocument.uri
        });
        const data = JSON.stringify({
            id: "text",
            json_string: json_data
        });
    // let ret = await getComdData(data);
    // let res_items: string[] = JSON.parse(ret).items;
    // let comlItems:CompletionItem[] = res_items.map(item => {
    //     return     {
    //         label: item,
    //         insertText: item,
    //         kind: CompletionItemKind.Text
    //     };
    // });
    // return comlItems;
    return [
    {
        label: 'TypeScript',
        kind: CompletionItemKind.Text,
        data: 1
    },
    {
        label: 'JavaScript',
        kind: CompletionItemKind.Text,
        data: 2
    }
    ];
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

// Make the text document manager listen on the connection
// for open, change and close text document events
documents.listen(connection);

// Listen on the connection
connection.listen();