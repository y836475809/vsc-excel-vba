import { spawn } from "child_process";
import * as fs from "fs";
import { LPSRequest, MakeReqData } from "./lsp-request";

export class VBALanguageServerUtil {
	lspRequest: LPSRequest;

	constructor(port: number){
		this.lspRequest = new LPSRequest(port);
	}

	async shutdown(): Promise<void>{
		try {
			await this.lspRequest.send(MakeReqData.shutdown());
			await this.wait("shutdown");
		} catch (error) {
			// 
		}
	}

	async launch(port: number, exeFilePath: string){
		if(await this.isReady()){
			return;
		}

		if(!exeFilePath){
			throw new Error(`No setting value for server filepath`);
		}
		if(!fs.existsSync(exeFilePath)){
			throw new Error(`Not find ${exeFilePath}`);
		}
		const p = spawn("cmd.exe", ["/c", `${exeFilePath} ${port}`], { detached: true });
		p.on("error", (error)=> {
			throw error;
		});
	
		await this.wait("ready");
	}

	async reset(){
		// if (this.client && this.client.state === State.Running) {
		await this.lspRequest.send(MakeReqData.reset());
		// }
	}
	
	private async wait(state: "ready"|"shutdown"){
		const watiTimeMs = 500;
		const countMax = 10 * 1000 / watiTimeMs;
		let waitCount = 0;
		while(true){
			if(waitCount > countMax){
				throw new Error(`Timed out waiting for server ${state}`);
			}
			waitCount++;
			await new Promise(resolve => {
				setTimeout(resolve, watiTimeMs);
			});
			try {
				await this.lspRequest.send(MakeReqData.isReady());
				if(state === "ready"){
					break;
				}
			} catch (error) {
				if(state === "shutdown"){
					break;
				}
			}
		}
	}
	
	private async isReady(): Promise<boolean>{
		try {
			await this.lspRequest.send(MakeReqData.isReady());
			return true;
		} catch (error) {
			return false;
		}
	}
}