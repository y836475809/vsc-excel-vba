namespace Hoge {
    export type CompletionItem = {
        DisplayText: string;
        CompletionText: string;
        Description: string;
        ReturnType: string;
        Kind: string;
    };

    export type CompletionItems = {
        items: CompletionItem[];
    };

    export type Location = {
        Positon: number;
        Line: number;
        Character: number;
    };
    export type DefinitionItem = {
        FilePath: string;
        Start: Location;
        End: Location;
    };
    export type DefinitionItems = {
        items: DefinitionItem[];
    };

    type CommandId = 
        "AddDocuments" 
        | "DeleteDocuments" 
        | "RenameDocument" 
        | "ChangeDocument" 
        | "Completion" 
        | "Definition" 
        | "Hover" 
        | "Shutdown"
        | "Reset"
        | "IgnoreShutdown"
        | "Debug:GetDocuments";
    export type Command = {
        Id: CommandId;
        FilePaths: string[];
        Position: number;
        Text: string;
    };

    export type RequestMethod = 
        "createFiles"
        | "deleteFiles"
        | "renameFiles"
        | "changeText";
    export type RequestRenameParam = {
        oldUri: string,
        newUri: string
    };
}
