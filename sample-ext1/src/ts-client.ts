import * as http from "http";

const options: http.RequestOptions = {
    port: 9088,
    method: "POST",
    headers: {
        "Content-Type": "application/json",
    },
};
const url = "http://localhost";

export function getComdData(send_data: any): Promise<string> {
    if(process.env.DUMMY_SERVER){
        return new Promise((resolve, reject) => {
            const jsonData = JSON.parse(send_data);
            if(jsonData.Id === "AddDocuments"){
                const data = JSON.stringify({
                    FilePaths: ['test1', 'test2'],
                    Texts: ['test1 text', 'test2 text']
                });
                resolve(data);
                return;
            }
            if(jsonData.Id === "Hover"){
                const data = JSON.stringify({
                    items:[{
                        DisplayText: 'test1',
                        Description: 'test1 text'
                    }]
                });
                resolve(data);
                return;
            }
            if(jsonData.Id === "Completion"){
                const data = JSON.stringify({
                    items:[{
                        DisplayText: "test1(val1)",
                        CompletionText: "test1",
                        Kind: "Method"
                    }]
                });
                resolve(data);
                return;
            }
            if(jsonData.Id === "Definition"){
                const data = JSON.stringify({
                    items:[{
                        FilePath: "dummy",
                        Start: 0,
                        End: 0
                    }]
                });
                resolve(data);
                return;
            }
            const data = JSON.stringify({
                items: ['test1', 'test2']
            });
            resolve(data);
            return;
        });
    }
    return new Promise((resolve, reject) => {
        const req = http.request(url, options, (res: http.IncomingMessage) => {
            let data = "";
            res.setEncoding('utf8');
            res.on('data', (chunk) => {
                data += chunk;
                // console.log('data');
            });

            res.on('end', () => {
                console.log("data=", data);
                console.log('end');
                resolve(data);
            });
        });
        req.on('error', function(e) {
            console.log('problem with request: ' + e.message);
            reject();
        });
        // const json_data = JSON.stringify({
        //     cmd: "OK",
        //     line: 10,
        //     col:25
        // });
        // const data = JSON.stringify({
        //     id: "text",
        //     json_string: json_data
        // });
        req.write(send_data);
        req.end();    
    });
}

// exports.getComdData = getComdData;