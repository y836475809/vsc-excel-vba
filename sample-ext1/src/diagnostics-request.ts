import { Diagnostic, DiagnosticSeverity } from "vscode-languageserver/node";
import { TextDocument } from 'vscode-languageserver-textdocument';
import { LPSRequest } from "./lsp-request";

const severityMap = new Map<string, DiagnosticSeverity>([
    ["Hidden", DiagnosticSeverity.Error],
    ["Info", DiagnosticSeverity.Information],
    ["Warning", DiagnosticSeverity.Warning],
    ["Error", DiagnosticSeverity.Error]
]);

export async function diagnosticsRequest(doc: TextDocument , fsPath: string, lpsRequest: LPSRequest){
    const data = {
    id: "Diagnostic",
        filepaths: [fsPath],
        position: 0,
        text: ""
    } as Hoge.Command;
    const items = await lpsRequest.send(data) as Hoge.DiagnosticItem[];
    const diagnostics = items.map(item => {
        const severity = severityMap.get(item.severity)!;
        const diagnostic: Diagnostic = {
            severity: severity,
            range: {
                start: doc.positionAt(item.start),
                end: doc.positionAt(item.end)
            },
            message: item.message,
            source: "server",
        };
        return diagnostic;
    });
    return diagnostics;
}