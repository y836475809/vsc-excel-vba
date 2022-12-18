import {
    createConnection,
    Diagnostic,
    DiagnosticSeverity,
    InitializeResult,
    ProposedFeatures,
    Range,
    TextDocuments,
    TextDocumentSyncKind,
} from 'vscode-languageserver/node';
import { TextDocument } from 'vscode-languageserver-textdocument';

const connection = createConnection();
connection.console.info(`Sample server running in node ${process.version}`);