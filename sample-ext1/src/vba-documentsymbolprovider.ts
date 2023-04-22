import * as vscode from 'vscode';

export class VbaDocumentSymbolProvider implements vscode.DocumentSymbolProvider {
    private getRootSymbol(document: vscode.TextDocument): vscode.DocumentSymbol{
        let symKind = vscode.SymbolKind.Module;
        let line = "Root";
        const fname = document.fileName;
        if(fname.endsWith(".bas") && document.lineCount > 0){
            line = document.lineAt(0).text;
            symKind = vscode.SymbolKind.Module;
        }
        if(fname.endsWith(".cls") && document.lineCount > 4){
            line = document.lineAt(4).text;
            symKind = vscode.SymbolKind.Class;
        }
        const rootName = line.replace(/Attribute\s+VB_Name\s+=\s+/, "").replace(/"/g, "");
        const range1 = document.lineAt(0).range;
        const range2 = document.lineAt(document.lineCount-1).range;
        const range = new vscode.Range(range1.start, range2.start);
        const symbol = new vscode.DocumentSymbol(
            rootName,
            "",
            symKind,
            range,
            range);
        return symbol;
    }

    public provideDocumentSymbols(document: vscode.TextDocument,
            token: vscode.CancellationToken): Thenable<vscode.DocumentSymbol[]> {
        return new Promise((resolve, reject) => {
            const regFuncStart = /^\s*(Private\s+|Public\s+){0,1}(Sub|Function)\s+\S+/i;
            const regFuncEnd = /^\s*(End)\s+(Sub|Function)\s*/i;
            const regVar = /^\s*(Private\s+|Public\s+){0,1}\S+\s+As\s+\S+/i;
            const symbols:vscode.DocumentSymbol[] = [];

            if(document.lineCount === 0){
                resolve(symbols);
            }
            const rootSymbol = this.getRootSymbol(document);
            rootSymbol.children = symbols;

            let varSymKind = vscode.SymbolKind.Variable;
            let funcSymKind = vscode.SymbolKind.Function;
            if(rootSymbol.kind === vscode.SymbolKind.Class){
                varSymKind = vscode.SymbolKind.Field;
                funcSymKind = vscode.SymbolKind.Method;
            }
            for (let i = 0; i < document.lineCount; i++) {
                const line = document.lineAt(i);
                if (regVar.test(line.text)) {
                    symbols.push(new vscode.DocumentSymbol(
                        line.text,
                        "",
                        varSymKind,
                        line.range,
                        line.range));
                }
                if (regFuncStart.test(line.text)) {
                    while(i < document.lineCount){
                        if (regFuncEnd.test(document.lineAt(i).text)) {
                            break;
                        }
                        i++;
                    }
                    if(i >= document.lineCount){
                        break;
                    }
                    const range1 = line.range;
                    const range2 = document.lineAt(i).range;
                    const range = new vscode.Range(range1.start, range2.start);
                    symbols.push(new vscode.DocumentSymbol(
                        line.text,
                        "",
                        funcSymKind,
                        range,
                        range));
                }
            }
            resolve([rootSymbol]); 
        });
    }
}