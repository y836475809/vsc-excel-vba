import * as http from "http";
import { URI } from 'vscode-uri';

const url = "http://localhost";

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

    send(json: Hoge.Command): Promise<any> {
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
    static addDocuments(uris: string[]): Hoge.Command {
        const filePaths = uris.map(uri => URI.parse(uri).fsPath);
        const param = {
            id: "AddDocuments",
            filepaths: filePaths,
            line: 0,
            chara: 0,
            text: ""
        } as Hoge.Command;
        return param;
    }

    static deleteDocuments(uris: string[]): Hoge.Command {
        const fsPaths = uris.map(uri => {
            return URI.parse(uri).fsPath;
        });
        const param = {
            id: "DeleteDocuments",
            filepaths: fsPaths,
            line: 0,
            chara: 0,
            text: ""
        } as Hoge.Command;  
        return param;
    }

    static renameDocuments(renameArgs: Hoge.RequestRenameParam[]): Hoge.Command[] {
        if(!renameArgs){
            return [];
        }
        const params: Hoge.Command[] = [];
        for(const renameArg of renameArgs){
            const oldUri = renameArg.olduri;
            const newUri = renameArg.newuri;
            const oldFsPath = URI.parse(oldUri).fsPath;
            const newFsPath = URI.parse(newUri).fsPath;
            const data = {
                id: "RenameDocument",
                filepaths: [oldFsPath, newFsPath],
                line: 0,
                chara: 0,
                text: ""
            } as Hoge.Command;  
            params.push(data);
        }
        return params;
    }

    static changeDocument(uri: string, text: string): Hoge.Command {
        const fsPath = URI.parse(uri).fsPath;
        const param = {
            id: "ChangeDocument",
            filepaths: [fsPath],
            position: 0,
            line: 0,
            chara: 0,
            text: text
        } as Hoge.Command;
        return param;
    }

    static diagnostics(fsPath: string): Hoge.Command {
        const param = {
            id: "Diagnostic",
            filepaths: [fsPath],
            line: 0,
            chara: 0,
            text: ""
            } as Hoge.Command;
        return param;
    }

    static reset(): Hoge.Command {
        const param = {
            id: "Reset",
            filepaths: [],
            line: 0,
            chara: 0,
            text: ""
        } as Hoge.Command;  
        return param;
    }

    static isReady(): Hoge.Command {
        const param = {
            id: "IsReady",
            filepaths: [],
            line: 0,
            chara: 0,
            text: ""
        } as Hoge.Command;
        return param;
    }

    static shutdown(): Hoge.Command {
        const param = {
            id: "Shutdown",
            filepaths: [],
            line: 0,
            chara: 0,
            text: ""
        } as Hoge.Command;
        return param;
    }
}