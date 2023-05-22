import * as path from "path";
import * as vscode from "vscode";
import { spawn } from "child_process";
import { Project } from "./project";

type VBAResponse = {
    code: string; // ok, error
    message: string;
    data: string; // json(str)
};

const basLineOffset = 1;
const clsLineOffset = 9;

export class VBACommands {
    private cmd: string;
    private scriptDirPath: string;
    private xlsmFileName: string;
    constructor(scriptDirPath: string){
        this.cmd = "powershell";
        this.scriptDirPath = scriptDirPath;
        this.xlsmFileName = "";
    }

    async exceue(project: Project, cmd: string){
        try {
            this.xlsmFileName = project.projectData.targetfilename;
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
                await this.resetBreakpoints();
                await this.gotoVBA();     
            }
            if(cmd === "export"){
                await this.export(project.srcDir);
            }
            if(cmd === "compile"){
                await this.compile();
            }
            if(cmd === "resetBreakpoints"){
                await this.resetBreakpoints();
            }
            vscode.window.showInformationMessage(`Successfully ${cmd}`);
        } catch (error: unknown) {
            let msg = `Failed to ${cmd}`;
            if(error instanceof Error){
                msg = `${msg}, ${error.message}`;
            }
            vscode.window.showErrorMessage(msg);
        }
    }

    async gotoVSCode(uris: string[]){
        const ret = await this.run("goto-vscode.ps1", []);
        const selJson = JSON.parse(ret.data);
        const moduleName = selJson["module_name"];
        const moduleType = selJson["module_type"];
        const startline = selJson["startline"] - 1;
        const startcol = selJson["startcol"] - 1;
        const endline = selJson["endline"] - 1;
        const endcol = selJson["endcol"] - 1;

        let lineoffset = 0;
        if(moduleType === "bas"){
            lineoffset = basLineOffset;
        }
        if(moduleType === "cls"){
            lineoffset = clsLineOffset;
        }

        const files = uris.filter(x => {
            return x.endsWith(`${moduleName}.${moduleType}`);
        });
        if(files.length > 0){
            const u = vscode.Uri.parse(files[0]);
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
        const modulename = fsPath.name;
        const ext = fsPath.ext;
        let line = editor.selection.start.line + 1;
        if(ext === ".bas"){
            line -= basLineOffset;
        }
        if(ext === ".cls"){
            line -= clsLineOffset;
        }
        await this.run("goto-vba.ps1", [modulename, line.toString()]);
    }

    async import(srcDir: string){
        await this.run("import.ps1", [`"${srcDir}"`]);
    }

    async export(distDir: string){
        await this.run("export.ps1", [`"${distDir}"`]);
    }

    async compile(){
        await this.run("compile.ps1", []);
    }

    async resetBreakpoints(){
        const dict = new Map<string, number[]>();
       
        vscode.debug.breakpoints.forEach(x => {
            const s = x as vscode.SourceBreakpoint;
            const loc = s.location;
            const ret = this.getVBALine(loc.uri, loc.range.start.line);
            if(ret){
                const modulename = ret[0];
                const line = ret[1];
                if(!dict.has(modulename)){
                    dict.set(modulename, []);
                }
                dict.get(modulename)?.push(line);
            }
        });
        const args: string[] = [];
        // m1:1-2-3 m2:2
        for (const [key, value] of dict) {
            args.push(`${key}:${value.join("-")}`);
        }
        await this.run("resetbreakpoints.ps1", args);
    }

    async runVBASubProc(xlsmFileName: string, procName: string){
        this.xlsmFileName = xlsmFileName;
        await this.run("run-vba-proc.ps1", [procName]);
    }

    private async run(scriptFileName: string, args: string[]): Promise<VBAResponse> {
        const scriptFilePath = path.join(this.scriptDirPath, scriptFileName);
        const ret = await this.spawnAsync(this.cmd, 
            ["-NoProfile", "-ExecutionPolicy", "Unrestricted", 
            `"${scriptFilePath}"`, this.xlsmFileName, ...args]);
        
        const res: VBAResponse = JSON.parse(ret);
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
        const modulename = fsPath.name;
        const ext = fsPath.ext;
        let line = vscodeLine + 1;
        if(ext === ".bas"){
            line -= basLineOffset;
        }
        if(ext === ".cls"){
            line -= clsLineOffset;
        }
        return [modulename, line];
    }
}