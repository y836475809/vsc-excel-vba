namespace Hoge {
    export type CompletionItem = {
        displaytext: string;
        completiontext: string;
        description: string;
        returntype: string;
        kind: string;
    };

    export type Location = {
        positon: number;
        line: number;
        character: number;
    };
    export type DefinitionItem = {
        filepath: string;
        start: Location;
        end: Location;
    };

    type Severity = 
        "Hidden"
        | "Info"
        | "Warning"
        | "Error";
    export type DiagnosticItem = {
        severity: Severity;
        message: string;
        startline: number;
        startchara: number;
        endline: number;
        endchara: number;
    };

    type CommandId = 
        "AddDocuments" 
        | "DeleteDocuments" 
        | "RenameDocument" 
        | "ChangeDocument" 
        | "Completion" 
        | "Definition" 
        | "Hover" 
        | "Diagnostic"
        | "Shutdown"
        | "Reset"
        | "IgnoreShutdown"
        | "Debug:GetDocuments";
    export type Command = {
        id: CommandId;
        filepaths: string[];
        line: number;
        chara: number;
        text: string;
    };

    export type RequestMethod = 
        "createFiles"
        | "deleteFiles"
        | "renameFiles"
        | "changeText";
    export type RequestRenameParam = {
        olduri: string,
        newuri: string
    };
}
