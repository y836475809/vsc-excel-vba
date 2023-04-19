import * as vscode from "vscode";
import * as path from "path";
import * as fs from "fs";

 

export class Project {
    projectData: Hoge.ProjectData; 
    projectFileName: string;
    constructor(projectFileName: string){
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
        await config.update("files.autoGuessEncoding", true);
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

    async readProject(dir: string): Promise<void> {
        const projectFp = path.join(dir,this.projectFileName);
        const json = await fs.promises.readFile(projectFp);
        this.projectData = JSON.parse(json.toString()) as Hoge.ProjectData;
    }

    async getSrcFileUris(): Promise<string[]> {
        const dir = this.projectData.srcdir;
        const listFiles = (dir: string): string[] =>
        fs.readdirSync(dir, { withFileTypes: true }).flatMap(dirent =>
            dirent.isFile() ? [`${dir}/${dirent.name}`] : listFiles(`${dir}/${dirent.name}`));
        const files = listFiles(dir);
        const srcFiles = files.filter(x => {
            return x.endsWith(".cls") || x.endsWith(".bas");
        });
        const uris = srcFiles.map(fp => vscode.Uri.file(fp).toString());
        return uris;
    }

    hasProject(wsPath: string | undefined): boolean{
        if(!wsPath){
            return false;
        }
        const pfs = path.join(wsPath, this.projectFileName);
        return fs.existsSync(pfs);
    }
}