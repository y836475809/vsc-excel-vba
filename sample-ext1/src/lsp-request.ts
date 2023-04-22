import * as http from "http";

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
