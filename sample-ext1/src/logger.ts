
type LogFunc = (msg: string) => void;

export class Logger {
    func: LogFunc;
    constructor(func: LogFunc){
        this.func = func;
    }

    info(msg: string){
        this.log("info", msg);
    }

    error(msg: string){
        this.log("error", msg);
    }

    private log(type: string, msg: string){
        const date = new Date();
        this.func(`${date.toLocaleString()}:${date.getMilliseconds()} [${type}] ${msg}`);
    }
}