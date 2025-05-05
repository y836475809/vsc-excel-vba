
import * as vscode from 'vscode';
import { VBACommands } from './vba-commands';
import { Project } from './project';
import { VBACreateFile } from "./vba-create-file";
import { SheetTreeDataProvider } from "./sheet-treedata-provider";


let sheetTDProvider: SheetTreeDataProvider;

export function vbaRegister(context: vscode.ExtensionContext, project: Project, vbaCommand: VBACommands) {
    sheetTDProvider = new SheetTreeDataProvider("sheetView");
    context.subscriptions.push(sheetTDProvider);
    
    vscode.commands.executeCommand("setContext", "vsc-excel-vba.showSheetView", true);

    context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.gotoVBA", async () => {
        await vbaCommand.exceue(project, "gotoVBA");
    }));
    context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.gotoVSCode", async () => {
        await vbaCommand.exceue(project, "gotoVSCode");
    }));
    context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.import", async () => {
        await vbaCommand.exceue(project, "importAndCompile");
    }));
    context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.export", async () => {
        if(await vscode.window.showInformationMessage(
            `Export?`, "Yes", "No") === "No"){
            return;
        }
        if(project.existSrcDir()){
            if(await vscode.window.showInformationMessage(
                `${project.srcDir} exists. Overwrite?`, "Yes", "No") === "No"){
                return;
            }
        }
        if(await vbaCommand.exceue(project, "export")){
            vscode.window.showInformationMessage(`Success export to ${project.srcDir}`);
        }
    }));
    context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.compile", async () => {
        await vbaCommand.exceue(project, "compile");
    }));
    context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.runVBAProc", async () => {
        await vbaCommand.exceue(project, "runVBAProc");
    }));
    context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.setBreakpoints", async () => {
        await vbaCommand.exceue(project, "setBreakpoints");
    }));

    context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.openSheet", async (args) => {	
        const xlsmFileName = project.projectData.excelfilename;
        await vbaCommand.openSheet(xlsmFileName, args.sheetName);
    }));
    context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.getSheetNames", async (args) => {	
        const xlsmFileName = project.projectData.excelfilename;
        const sheetNames = await vbaCommand.getSheetNames(xlsmFileName);
        sheetTDProvider.refresh(sheetNames);
    }));
    context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.newClassFile", async (uri: vscode.Uri) => {	
        const vbacf = new VBACreateFile();
        await vbacf.create(uri, "class");
    }));
    context.subscriptions.push(vscode.commands.registerCommand("vsc-excel-vba.newModuleFile", async (uri: vscode.Uri) => {	
        const vbacf = new VBACreateFile();
        await vbacf.create(uri, "module");
    }));
};
