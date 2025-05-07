import * as vscode from 'vscode';
import { Commands } from './commands';
import { Project } from './project';
import { CreateFile } from "./create-file";
import { SheetTreeDataProvider } from "./sheet-treedata-provider";


let sheetTDProvider: SheetTreeDataProvider;

async function checkExport(project: Project): Promise<boolean> {
    if(await vscode.window.showInformationMessage(
        `Export?`, "Yes", "No") === "No"){
        return false;
    }
    if(project.existSrcDir()){
        if(await vscode.window.showInformationMessage(
            `${project.srcDir} exists. Overwrite?`, "Yes", "No") === "No"){
            return false;
        }
    }
    return true;
}

export function register(context: vscode.ExtensionContext, project: Project, commands: Commands) {
    sheetTDProvider = new SheetTreeDataProvider("sheetView");
    context.subscriptions.push(sheetTDProvider);
    
    vscode.commands.executeCommand("setContext", "vsc-excel-vba.showSheetView", true);

    context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.gotoVBA", async () => {
        await commands.exceue(project, "gotoVBA");
    }));
    context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.gotoVSCode", async () => {
        await commands.exceue(project, "gotoVSCode");
    }));
    context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.import", async () => {
        await commands.exceue(project, "importAndCompile");
    }));
    context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.export", async () => {
        if(!await checkExport(project)){
            return;
        }
        if(await commands.exceue(project, "export")){
            vscode.window.showInformationMessage(`Success export to ${project.srcDir}`);
        }
    }));
    context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.openExport", async () => {
        if(!await checkExport(project)){
            return;
        }
        vscode.window.withProgress({
            location: vscode.ProgressLocation.Notification, 
            title: "Export files"
        }, async progress => {
            return new Promise<void>(async (resolve)=>{
                if(await commands.exceue(project, "openExport")){
                    vscode.window.showInformationMessage(`Success export to ${project.srcDir}`);
                }
                resolve();
            });
        });
    }));
    context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.compile", async () => {
        await commands.exceue(project, "compile");
    }));
    context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.runVBAProc", async () => {
        await commands.exceue(project, "runVBAProc");
    }));
    context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.setBreakpoints", async () => {
        await commands.exceue(project, "setBreakpoints");
    }));

    context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.openSheet", async (args) => {	
        const xlsmFileName = project.projectData.excelfilename;
        await commands.openSheet(xlsmFileName, args.sheetName);
    }));
    context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.getSheetNames", async (args) => {	
        const xlsmFileName = project.projectData.excelfilename;
        const sheetNames = await commands.getSheetNames(xlsmFileName);
        sheetTDProvider.refresh(sheetNames);
    }));
    context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.newClassFile", async (uri: vscode.Uri) => {	
        const cf = new CreateFile();
        await cf.create(uri, "class");
    }));
    context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.newModuleFile", async (uri: vscode.Uri) => {	
        const cf = new CreateFile();
        await cf.create(uri, "module");
    }));
};
