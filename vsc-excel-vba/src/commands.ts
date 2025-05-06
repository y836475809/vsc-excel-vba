import * as path from "path";
import * as vscode from "vscode";
import { spawn } from "child_process";
import { Project } from "./project";

type Response = {
    code: string; // ok, error
    message: string;
    data: string; // json(str)
};

const basLineOffset = 1;
const clsLineOffset = 9;

export class Commands {
    private cmd: string;
    private scriptDirPath: string;
    private xlsmFileName: string;
    constructor(scriptDirPath: string){
        this.cmd = "powershell";
        this.scriptDirPath = scriptDirPath;
        this.xlsmFileName = "";
    }

    async tryCatch(cmd: string, func: ()=>Promise<void>): Promise<boolean>{
        try {
            await func();
            return true;
        } catch (error: unknown) {
            let msg = `Failed to ${cmd}`;
            if(error instanceof Error){
                msg = `${msg}, ${error.message}`;
            }
            vscode.window.showErrorMessage(msg);
            vscode.window.showErrorMessage(
                `Open excel file or Enable 'Trust Access to Visual Basic Project'
                (File > Options > Trust Center > Trust Center Settings > Macro Settings)`);
            return false;
        } 
    }

    async exceue(project: Project, cmd: string): Promise<boolean>{
        return await this.tryCatch(cmd, async () => {
            this.xlsmFileName = project.projectData.excelfilename;
            if(!this.xlsmFileName){
                throw new Error(
                    "Excel file name is empty");
            }
    
            if(cmd === "gotoVSCode"){
                await this.gotoVSCode(await project.getSrcFileUris());
            }
            if(cmd === "gotoVBA"){
                await this.gotoVBA();
            }
            if(cmd === "import"){
                await this.import(project.srcDir);
                await this.setBreakpoints();
                await this.gotoVBA();     
            }
            if(cmd === "importAndCompile"){
                await this.import(project.srcDir);
                await this.setBreakpoints();
                await this.gotoVBA();
                await this.compile();
            }
            if(cmd === "export"){
                await this.export(project.srcDir);
            }
            if(cmd === "compile"){
                await this.compile();
            }
            if(cmd === "setBreakpoints"){
                await this.setBreakpoints();
            }
            if(cmd === "runVBAProc"){
                await this.runVBAProc();
            }
        });
    }

    async gotoVSCode(uris: vscode.Uri[]){
        const ret = await this.run("goto-vscode.ps1", []);
        const selJson = JSON.parse(ret.data);
        const objectName = selJson["objectname"];
        const objectType = selJson["objecttype"];
        const startline = selJson["startline"] - 1;
        const startcol = selJson["startcol"] - 1;
        const endline = selJson["endline"] - 1;
        const endcol = selJson["endcol"] - 1;

        let lineoffset = 0;
        if(objectType === "bas"){
            lineoffset = basLineOffset;
        }
        if(objectType === "cls"){
            lineoffset = clsLineOffset;
        }

        const encFname = encodeURIComponent(`${objectName}.${objectType}`);
        const files = uris.filter(x => {
            const fname = x.toString().split("/").at(-1);
            return fname === encFname;
        });
        if(files.length > 0){
            const u = files[0];
            const doc = await vscode.workspace.openTextDocument(u);
            if(!doc){
                return;
            }
            const editor = await vscode.window.showTextDocument(doc, vscode.ViewColumn.Active, false);
            if(!editor){
                return;
            }
            const vsStartline = startline + lineoffset;
            const vsEndline = endline + lineoffset;
            editor.selection = new vscode.Selection(
                vsStartline, startcol,
                vsEndline, endcol);
            const range = editor.document.lineAt(vsStartline).range;
            editor.revealRange(range, 
                vscode.TextEditorRevealType.InCenterIfOutsideViewport);
        }
    }

    async gotoVBA(){
        const editor = vscode.window.activeTextEditor;
        if(!editor){
            return;
        }
        const fsPath = path.parse(editor.document.uri.fsPath);
        const objectName = fsPath.name;
        const ext = fsPath.ext;
        let line = editor.selection.start.line + 1;
        if(ext === ".bas"){
            line -= basLineOffset;
        }
        if(ext === ".cls"){
            line -= clsLineOffset;
        }
        await this.run("goto-vba.ps1", [objectName, line.toString()]);
    }

    async import(srcDir: string){
        await this.run("import.ps1", [`'${srcDir}'`]);
    }

    async export(distDir: string){
        await this.run("export.ps1", [`'${distDir}'`]);
    }

    async compile(){
        await this.run("compile.ps1", []);
    }

    async setBreakpoints(){
        const dict = new Map<string, number[]>();
       
        vscode.debug.breakpoints.forEach(x => {
            const s = x as vscode.SourceBreakpoint;
            const loc = s.location;
            const ret = this.getVBALine(loc.uri, loc.range.start.line);
            if(ret){
                const objectName = ret[0];
                const line = ret[1];
                if(!dict.has(objectName)){
                    dict.set(objectName, []);
                }
                dict.get(objectName)?.push(line);
            }
        });
        const args: string[] = [];
        // m1:1-2-3 m2:2
        for (const [key, value] of dict) {
            args.push(`${key}:${value.join("-")}`);
        }
        await this.run("set-breakpoints.ps1", args);
    }

    async runVBAProc(){
        await this.tryCatch("runVBAProc", async () => {
            const editor = vscode.window.activeTextEditor;
            if(!editor){
                return;
            }
            const sel = editor.selection;
            const uri = editor.document.uri;
            const symbols = await vscode.commands.executeCommand<vscode.DocumentSymbol[]>(
                "vscode.executeDocumentSymbolProvider", uri);
            if(!symbols?.length){
                return;
            }
            const targetSymbols = symbols[0].children.filter(x => {
                return (x.kind === vscode.SymbolKind.Function || x.kind === vscode.SymbolKind.Method) 
                    && x.range.start.line <= sel.start.line 
                    && sel.end.line <= x.range.end.line;
            });
            if(!targetSymbols.length){
                return;
            }
            const objectName = path.parse(uri.fsPath).name;
            const procName = targetSymbols[0].name;
            await this.run("run-vba-proc.ps1", [`${objectName}.${procName}`]);
        });
    }

    async openSheet(xlsmFileName: string, sheetName: string){
        await this.tryCatch("openSheet", async () => {
            this.xlsmFileName = xlsmFileName;
            await this.run("open-sheet.ps1", [sheetName]);
        });
    }

    async getSheetNames(xlsmFileName: string): Promise<string[]> {
        this.xlsmFileName = xlsmFileName;
        const ret = await this.run("get-sheet-names.ps1", []);
        const sheetNames: string[] = JSON.parse(ret.data);
        return sheetNames;
    }

    private async run(scriptFileName: string, args: string[]): Promise<Response> {
        const scriptFilePath = path.join(this.scriptDirPath, scriptFileName);
        const ret = await this.spawnAsync(this.cmd, 
            ["-NoProfile", "-ExecutionPolicy", "Unrestricted", 
            `"${scriptFilePath}"`, this.xlsmFileName, ...args]);
        
        const res: Response = JSON.parse(ret);
        if(res.code === "error"){
            const msg = res.message;
            throw new Error(msg);
        }   
        return res;
    }

    private spawnAsync(command: string, args: string[]): Promise<string> {
		return new Promise((resolve) => {
			const child = spawn(command, args, {
				windowsHide: true,
				shell: true,
				// detached: true
			});
	
			const result: string[] = [];
			child.stdout.on("data", (data) => {
				result.push(data);
			});
		
			child.on("close", () => {
				return resolve(result.join(""));
			});
		});
  	};

    private getVBALine(uri: vscode.Uri, vscodeLine: number) : [string, number] | undefined{
        const fsPath = path.parse(uri.fsPath);
        const objectName = fsPath.name;
        const ext = fsPath.ext;
        let line = vscodeLine + 1;
        if(ext === ".bas"){
            line -= basLineOffset;
        }
        if(ext === ".cls"){
            line -= clsLineOffset;
        }
        return [objectName, line];
    }
}