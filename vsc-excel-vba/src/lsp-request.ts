import * as http from "http";
import * as vscode from "vscode";

const url = "http://localhost";

export type RenameParam = {
    oldUri: vscode.Uri;
    newUri: vscode.Uri;
};

export class LPSRequest {
    private port: number;
    constructor(port: number){
        this.port = port;
    }

    private getOptions(data: string): http.RequestOptions {
        return {
            port: this.port,
            method: "POST",
            headers: {
                // eslint-disable-next-line @typescript-eslint/naming-convention
                "Content-Length": Buffer.byteLength(data),
                // eslint-disable-next-line @typescript-eslint/naming-convention
                "Content-Type": "application/json",
            } 
        };
    }

    send(json: VEV.RequestParam): Promise<any> {
        return new Promise((resolve, reject) => {
            const jsonStr = JSON.stringify(json);
            const options = this.getOptions(jsonStr);
            const req = http.request(url, options, (res: http.IncomingMessage) => {
                let data = "";
                res.setEncoding('utf8');
                res.on('data', (chunk) => {
                    data += chunk;
                });
                res.on('end', () => {
                    if(res.statusCode !== 200){
                        reject(new Error(`statusCode=${res.statusCode}, ${res.statusMessage}`));
                        return;
                    }
                    if(data.length === 0){
                        resolve({});
                    }else{
                        resolve(JSON.parse(data));
                    }  
                });
            });
            req.on('error', function(e) {
                reject(e);
            });
            req.write(jsonStr);
            req.end();    
        });
    }
}

export class MakeReqData {
    static addDocuments(uris: vscode.Uri[]): VEV.RequestParam {
        const filePaths = uris.map(uri => uri.fsPath);
        const param = {
            id: "AddDocuments",
            filepaths: filePaths,
            line: 0,
            chara: 0,
            text: ""
        } as VEV.RequestParam;
        return param;
    }

    static deleteDocuments(uris: vscode.Uri[]): VEV.RequestParam {
        const fsPaths = uris.map(uri => {
            return uri.fsPath;
        });
        const param = {
            id: "DeleteDocuments",
            filepaths: fsPaths,
            line: 0,
            chara: 0,
            text: ""
        } as VEV.RequestParam;  
        return param;
    }

    static renameDocuments(renameArgs: RenameParam[]): VEV.RequestParam[] {
        if(!renameArgs){
            return [];
        }
        const params: VEV.RequestParam[] = [];
        for(const renameArg of renameArgs){
            const oldFsPath = renameArg.oldUri.fsPath;
            const newFsPath = renameArg.newUri.fsPath;
            const data = {
                id: "RenameDocument",
                filepaths: [oldFsPath, newFsPath],
                line: 0,
                chara: 0,
                text: ""
            } as VEV.RequestParam;  
            params.push(data);
        }
        return params;
    }

    static changeDocument(uri: vscode.Uri, text: string): VEV.RequestParam {
        const fsPath = uri.fsPath;
        const param = {
            id: "ChangeDocument",
            filepaths: [fsPath],
            position: 0,
            line: 0,
            chara: 0,
            text: text
        } as VEV.RequestParam;
        return param;
    }

    static diagnostics(fsPath: string): VEV.RequestParam {
        const param = {
            id: "Diagnostic",
            filepaths: [fsPath],
            line: 0,
            chara: 0,
            text: ""
            } as VEV.RequestParam;
        return param;
    }

    static reset(): VEV.RequestParam {
        const param = {
            id: "Reset",
            filepaths: [],
            line: 0,
            chara: 0,
            text: ""
        } as VEV.RequestParam;  
        return param;
    }

    static isReady(): VEV.RequestParam {
        const param = {
            id: "IsReady",
            filepaths: [],
            line: 0,
            chara: 0,
            text: ""
        } as VEV.RequestParam;
        return param;
    }

    static shutdown(): VEV.RequestParam {
        const param = {
            id: "Shutdown",
            filepaths: [],
            line: 0,
            chara: 0,
            text: ""
        } as VEV.RequestParam;
        return param;
    }
}