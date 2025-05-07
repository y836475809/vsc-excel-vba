import * as vscode from "vscode";
import * as path from "path";
import * as fs from "fs";

const PROJECT_FILENAME: string = "vbaproject.json";
const SRC_DIRNAME: string = "src";

export class Project {
    srcDir: string;
    projectData: VEV.ProjectData; 
    excelFilePath: string;

    constructor(){
        this.srcDir = "";
        this.excelFilePath = "";
        this.projectData = {
            excelfilename:""
        };
    }

    get projectFileName(){
        return PROJECT_FILENAME;
    }

    async createProject(excelFilePath: string){
        this.excelFilePath = excelFilePath;
        await this.createSetting();
        const dirPath = path.dirname(excelFilePath);
		const filename = path.basename(excelFilePath);

		const projectFp = path.join(dirPath, PROJECT_FILENAME);
		const data: VEV.ProjectData = {
			excelfilename: filename
		};
		await fs.promises.writeFile(projectFp, JSON.stringify(data, null, 4));

        this.srcDir = path.join(dirPath, SRC_DIRNAME);
        if(!fs.existsSync(this.srcDir)){
            await fs.promises.mkdir(this.srcDir);
        }
    }

    async readProject(): Promise<void> {
        const wsPath = this.getWorkspacePath();
        if(!wsPath){
			const msg = `Not find ${PROJECT_FILENAME}`;
			throw new Error(msg);
		}
        const projectFp = path.join(wsPath, PROJECT_FILENAME);
        const json = await fs.promises.readFile(projectFp);
        this.projectData = JSON.parse(json.toString()) as VEV.ProjectData;
        this.excelFilePath = path.join(wsPath, this.projectData.excelfilename);
        this.srcDir = path.join(wsPath, SRC_DIRNAME);
        if(!fs.existsSync(this.srcDir)){
            await fs.promises.mkdir(this.srcDir);
        }
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

    hasProject(): boolean{
        const wsPath = this.getWorkspacePath();
        if(!wsPath){
            throw new Error(`Not find workspace Folders`);
        }
        const pfs = path.join(wsPath, PROJECT_FILENAME);
        return fs.existsSync(pfs);
    }

    existSrcDir(): boolean {
        return fs.existsSync(this.srcDir);
    }

    private async createSetting() {
        const config = vscode.workspace.getConfiguration();
        await config.update(
            "[vb]", { 
                "files.encoding": "shiftjis",
            }, false);
        await config.update(
            "files.autoGuessEncoding", true, false);
        await config.update(
            "files.associations", {
                "*.bas": "vba",
                "*.cls": "vba"
            }, false);
    }

    private getWorkspacePath(): string | undefined{
        const wf = vscode.workspace.workspaceFolders;
        return (wf && (wf.length > 0)) ? wf[0].uri.fsPath : undefined;
    }
}