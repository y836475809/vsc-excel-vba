import * as http from "http";

const url = "http://localhost";

export class LPSRequest {
    private options: http.RequestOptions;

    constructor(port: number){
        this.options = {
            port: port,
            method: "POST",
            headers: {
                "Content-Type": "application/json",
            } 
        };
    }

    send(json: Hoge.Command): Promise<any> {
        return new Promise((resolve, reject) => {
            const req = http.request(url, this.options, (res: http.IncomingMessage) => {
                let data = "";
                res.setEncoding('utf8');
                res.on('data', (chunk) => {
                    data += chunk;
                });
                res.on('end', () => {
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
            req.write(JSON.stringify(json));
            req.end();    
        });
    }
}
