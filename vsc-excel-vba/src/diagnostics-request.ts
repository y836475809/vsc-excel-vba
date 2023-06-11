import * as vscode from "vscode";
import { LPSRequest, MakeReqData } from "./lsp-request";

const severityMap = new Map<string, vscode.DiagnosticSeverity>([
    ["Hidden",  vscode.DiagnosticSeverity.Error],
    ["Info",    vscode.DiagnosticSeverity.Information],
    ["Warning", vscode.DiagnosticSeverity.Warning],
    ["Error",   vscode.DiagnosticSeverity.Error]
]);

export async function diagnosticsRequest(fsPath: string, lpsRequest: LPSRequest){
    const data = MakeReqData.diagnostics(fsPath);
    const items = await lpsRequest.send(data) as Hoge.DiagnosticItem[];
    const diagnostics = items.map(item => {
        const severity = severityMap.get(item.severity)!;
        const diagnostic= new vscode.Diagnostic(
            new vscode.Range(
                item.startline, item.startchara,
                item.endline, item.endchara
            ),
            item.message,
            severity
        );
        return diagnostic;
    });
    return diagnostics;
}