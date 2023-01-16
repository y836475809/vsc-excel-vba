import * as http from "http";

const options: http.RequestOptions = {
    port: 9088,
    method: "POST",
    headers: {
        "Content-Type": "application/json",
    },
};
const url = "http://localhost";

export function getComdData(json: any): Promise<any> {
    return new Promise((resolve, reject) => {
        const req = http.request(url, options, (res: http.IncomingMessage) => {
            let data = "";
            res.setEncoding('utf8');
            res.on('data', (chunk) => {
                data += chunk;
            });

            res.on('end', () => {
                resolve(JSON.parse(data));
            });
        });
        req.on('error', function(e) {
            reject(e);
        });
        req.write(JSON.stringify(json));
        req.end();    
    });
}