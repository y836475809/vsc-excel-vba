import * as helper from "../helper";

export function run(): Promise<void> {
	return helper.run(__dirname)();
}
