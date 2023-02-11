import * as vscode from "vscode";

export class Project {
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
}