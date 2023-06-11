import * as vscode from "vscode";
import * as path from "path";
import * as fs from "fs";

 

export class Project {
    srcDir: string;
    projectData: Hoge.ProjectData; 
    projectFileName: string;
    constructor(projectFileName: string){
        this.srcDir = "";
        this.projectFileName = projectFileName;
        this.projectData = {
            targetfilename:"",
            srcdir: ""
        };
    }

    async setupConfig() {
        const config = vscode.workspace.getConfiguration();
        const editorConfig: any = await config.get("editor.quickSuggestions");
        if(editorConfig["other"] !== "off"){
            editorConfig["other"] = "on";
            await config.update(
                "editor.quickSuggestions",
                editorConfig,
                vscode.ConfigurationTarget.Workspace
            );
        }

        await config.update(
            "[vb]", { 
                "files.encoding": "shiftjis",
                "files.autoGuessEncoding": true
            }, false);
    }

    async createProject(targetFilePath: string){
        // const filename = targetFilePath.replace(wsPath, "");
		const filename = path.basename(targetFilePath);
		const projectFp = path.join(path.dirname(targetFilePath), this.projectFileName);
		const data: Hoge.ProjectData = {
			targetfilename: filename,
			srcdir: "src"
		};
		await fs.promises.writeFile(projectFp, JSON.stringify(data, null, 4));
    }

    async readProject(): Promise<void> {
        const wsPath = this.getWorkspacePath();
        if(!wsPath){
			const msg = `Not find ${this.projectFileName}`;
			throw new Error(msg);
		}
        const projectFp = path.join(wsPath,this.projectFileName);
        const json = await fs.promises.readFile(projectFp);
        this.projectData = JSON.parse(json.toString()) as Hoge.ProjectData;
        this.srcDir = path.join(wsPath,this.projectData.srcdir);
    }

    async getSrcFileUris(): Promise<vscode.Uri[]> {
        const listFiles = (dir: string): string[] =>
        fs.readdirSync(dir, { withFileTypes: true }).flatMap(dirent =>
            dirent.isFile() ? [`${dir}/${dirent.name}`] : listFiles(`${dir}/${dirent.name}`));
        const files = listFiles(this.srcDir);
        const srcFiles = files.filter(x => {
            return x.endsWith(".cls") || x.endsWith(".bas");
        });
        const uris = srcFiles.map(fp => vscode.Uri.file(fp));
        return uris;
    }

    getDefinitionFileUris(context: vscode.ExtensionContext): vscode.Uri[] {
        const dirPath = context.asAbsolutePath("d.vb");
        if(!fs.existsSync(dirPath)){
            return [];
        }
        const fsPaths = fs.readdirSync(dirPath, { withFileTypes: true })
        .filter(dirent => {
            return dirent.isFile() && (dirent.name.endsWith(".d.vb"));
        }).map(dirent => path.join(dirPath, dirent.name));
        const uris = fsPaths.map(fp => vscode.Uri.file(fp));
        return uris;
    }

    hasProject(): boolean{
        const wsPath = this.getWorkspacePath();
        if(!wsPath){
            return false;
        }
        const pfs = path.join(wsPath, this.projectFileName);
        return fs.existsSync(pfs);
    }

    private getWorkspacePath(): string | undefined{
        const wf = vscode.workspace.workspaceFolders;
        return (wf && (wf.length > 0)) ? wf[0].uri.fsPath : undefined;
    }
}