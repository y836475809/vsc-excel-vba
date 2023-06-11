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

    export type ReferencesItem = {
        filepath: string;
        start: Location;
        end: Location;
    };

    export type SignatureHelpArgItem = {
        name: string;
        astype: string;
    };
    export type SignatureHelpItem = {
        displaytext: string;
        description: string;
        returntype: string;
        kind: string;
        args: SignatureHelpArgItem[];
        activeparameter: number;
    };

    export type RequestId = 
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
        | "IsReady"
        | "IgnoreShutdown"
        | "References"
        | "SignatureHelp"
        | "Debug:GetDocuments";
    export type RequestParam = {
        id: RequestId;
        filepaths: string[];
        line: number;
        chara: number;
        text: string;
    };

    export type RequestRenameParam = {
        olduri: string,
        newuri: string
    };

    export type ProjectData = {
        targetfilename: string;
        srcdir: string;
    };
}
