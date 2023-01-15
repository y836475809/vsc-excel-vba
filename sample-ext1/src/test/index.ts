import * as path from 'path';
import * as helper from "./helper";

export function run(): Promise<void> {
	const dirName = process.env["MOCHA_TEST_DIRNAME"];
	return helper.run(path.resolve(__dirname, `./${dirName}`))();
}
