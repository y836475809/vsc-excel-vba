import * as vscode from "vscode";
import * as path from 'path';
import { VBAObjectNameValidation } from "./file-name-validation";

export class VbaAttributeError extends Error {
    line: number;
    endpos: number;
    constructor(message: string, line: number, endpos: number){
        super(message);
        this.line = line;
        this.endpos = endpos;
    }
}

export function makeAttributeDiagnostics(errors: VbaAttributeError[]){
    return errors.map(x => {
        const msg = x.message;
        const line = x.line;
        const endpos = x.endpos;
        const diagnostic = new vscode.Diagnostic(     
            new vscode.Range(
                line, 0,
                line, endpos
            ),
            msg,
            vscode.DiagnosticSeverity.Error
        );
        return diagnostic;
    });
}

export class VbaAttributeValidation {
    validate(uri: vscode.Uri, text: string){
        const fp = path.parse(uri.fsPath);
        const name = fp.name;
        const lines = text.split("\r").map(x=>{
            return x.replace("\n", "");
        });

        if(uri.fsPath.endsWith(".bas")){
            this.validateBAS(name, lines);
        }
        if(uri.fsPath.endsWith(".cls")){
            this.validateCLS(name, lines);
        }
    }

    private validateObjectName(line: string, lineNum: number, lineLen: number): VbaAttributeError[] {
        const errors: VbaAttributeError[] = [];

        const validate = new VBAObjectNameValidation();
        const objectName = line.split("=")[1].trim().replace(/^"|"$/g, "");
        if(!validate.prefix(objectName)){
            const msg = `VB_Name = "${objectName}" contains illegal prefix`;
            errors.push(new VbaAttributeError(msg, lineNum, lineLen)); 
        }
        if(!validate.symbol(objectName)){
            const msg = `VB_Name = "${objectName}" contains illegal symbols`;
            errors.push(new VbaAttributeError(msg, lineNum, lineLen)); 
        }
        if(!validate.len(objectName)){
            const msg = `VB_Name = "${objectName}" is too long`;
            errors.push(new VbaAttributeError(msg, lineNum, lineLen)); 
        }
        return errors;
    }

    private validateBAS(name: string, lines: string[]) {
        const errors: VbaAttributeError[] = [];

        if(lines.length < 1){
            const msg = "Not enough Attribute";
            errors.push(new VbaAttributeError(msg, 0, 0));
        }
        if(errors.length){
            throw errors;
        }

        const lineNum = 0;
        const len = lines[lineNum].length;
        const reg = new RegExp(`Attribute\\s+VB_Name\\s*=\\s*"${name}"`);
        if(!reg.test(lines[0])){
            const msg = `Correct is Attribute VB_Name = "${name}"`;
            errors.push(new VbaAttributeError(msg, lineNum, len-1));
        }
        errors.push(...this.validateObjectName(lines[lineNum], lineNum, len-1));
        
        if(errors.length){
            throw errors;
        }
    }

    private validateCLS(name: string, lines: string[]){
        const errors: VbaAttributeError[] = [];

        if(lines.length < 9){
            const msg = "Not enough Attribute";
            errors.push(new VbaAttributeError(msg, 0, 0));
        }
        if(errors.length){
            throw errors;
        }

        function getLineEndpos(linenum: number): [string, number]{
            const line = lines[linenum];
            const endpos = line.length>0?line.length - 1:0;
            return [line, endpos];
        }

        let linenum = 0;
        let [line, endpos] = getLineEndpos(linenum);
        if(!/VERSION\s+1.0\s+CLASS/.test(line)){
            const msg = `Correct is VERSION 1.0 CLASS`;
            errors.push(new VbaAttributeError(msg, linenum, endpos));
        }

        [line, endpos] = getLineEndpos(++linenum);
        if(line !== "BEGIN"){
            const msg = `Correct is BEGIN`;
            errors.push(new VbaAttributeError(msg, linenum, endpos));
        }

        [line, endpos] = getLineEndpos(++linenum);
        if(!/MultiUse\s+=\s+-1/.test(line)){
            const msg = `Correct is MultiUse = -1`;
            errors.push(new VbaAttributeError(msg, linenum, endpos));
        }

        [line, endpos] = getLineEndpos(++linenum);
        if(line !== "END"){
            const msg = `Correct is END`;
            errors.push(new VbaAttributeError(msg, linenum, endpos));
        }

        [line, endpos] = getLineEndpos(++linenum);
        const reg = new RegExp(`Attribute\\s+VB_Name\\s*=\\s*"${name}"`);
        if(!reg.test(line)){
            const msg = `Correct is Attribute VB_Name = "${name}"`;
            errors.push(new VbaAttributeError(msg, linenum, endpos));
        }
        errors.push(...this.validateObjectName(line, linenum, endpos));

        [line, endpos] = getLineEndpos(++linenum);
        if(!/Attribute\s+VB_GlobalNameSpace\s+=\s+(True|False)/.test(line)){
            const msg = `Correct is Attribute VB_GlobalNameSpace = True|False`;
            errors.push(new VbaAttributeError(msg, linenum, endpos));
        }

        [line, endpos] = getLineEndpos(++linenum);
        if(!/Attribute\s+VB_Creatable\s+=\s+(True|False)/.test(line)){
            const msg = `Correct is Attribute VB_Creatable = True|False`;
            errors.push(new VbaAttributeError(msg, linenum, endpos));
        }

        [line, endpos] = getLineEndpos(++linenum);
        if(!/Attribute\s+VB_PredeclaredId\s+=\s+(True|False)/.test(line)){
            const msg = `Correct is Attribute VB_PredeclaredId = True|False`;
            errors.push(new VbaAttributeError(msg, linenum, endpos));
        }

        [line, endpos] = getLineEndpos(++linenum);
        if(!/Attribute\s+VB_Exposed\s+=\s+(True|False)/.test(line)){
            const msg = `Correct is Attribute VB_Exposed = True|False`;
            errors.push(new VbaAttributeError(msg, linenum, endpos));
        }

        if(errors.length){
            throw errors;
        }
    }
}