import * as http from "http";

type HttpServer = http.Server<typeof http.IncomingMessage, typeof http.ServerResponse>;

export class LspMockServer {
    srever: HttpServer;
    callBackAddDocuments:(json: any) => any;
    constructor(){
        this.callBackAddDocuments = (json: any) => {
            return {};
        };
        this.srever = http.createServer((req, res) => {
            if(req.method?.toLowerCase() === "get") {
                console.log(`get mock server`);
                res.writeHead(404);
                res.end();
                return;
            }
            let body = "";
            req.on("data", (chunk: string) => {
                body += chunk;
            });
            req.on("end", () => {
                let resJson = "";
                const json = JSON.parse(body);
                if(json.Id === "Shutdown"){  
                    req.destroy();
                    this.srever.close();
                }else{
                    resJson = this.callBackAddDocuments(json);
                }
                
                const head = {
                    "Content-Type": "json"
                };
                res.writeHead(202, head);
                res.end(JSON.stringify(resJson));
            });
        });
    }

    listen(port: number){
        this.srever.listen(port);
        console.log(`start mock server port=${port}`);
    }
    
    close(){
        this.srever.close();
        console.log(`close mock server`);
    }
}

const argv = process.argv;
if(argv.length >= 3 && argv[2] === "start"){
    const resMap = new Map<string, any>([
        ["Completion", { items: [
            {
                DisplayText: "test1(val1)",
                CompletionText: "test1",
                Kind: "Method"
            }
        ]}],
        ["Definition", { items: [
            {
                FilePath: "m1.bas",
                Start: { Line: 0, Character: 0 },
                End: { Line: 0, Character: 0 },
            }
        ]}],
        ["Hover", { items: [
            {
                DisplayText: "test1(val1)",
                Description: "Description test1",
            }
        ]}],
    ]);
    const mockServer = new LspMockServer();
    mockServer.callBackAddDocuments = (json: any) => {
        if(resMap.has(json.Id)){
            return resMap.get(json.Id);
        }
        return {};
    };
    mockServer.listen(9088);
}