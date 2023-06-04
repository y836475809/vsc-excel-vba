import * as path from 'path';
import * as os from "os";
import { runTests } from '@vscode/test-electron';
import * as glob from 'glob';

async function main() {
	const testsRoot = path.resolve(__dirname, '.');
	glob('e2e_*/', { cwd: testsRoot }, async (err, files) => {
		if (err) {
			throw err;
		}
		for (const dirName of files) {
			try {
				// The folder containing the Extension Manifest package.json
				// Passed to `--extensionDevelopmentPath`
				const extensionDevelopmentPath = path.resolve(__dirname, '../../');
		
				// The path to test runner
				// Passed to --extensionTestsPath
				const extensionTestsPath = path.resolve(__dirname, `./index`);
		
				const testWorkspace = path.resolve(__dirname, `../../src/test/${dirName}/fixture`);

				process.env["MOCHA_TEST_DIRNAME"] = dirName;
				// Download VS Code, unzip it and run the integration test
				// await runTests({ extensionDevelopmentPath, extensionTestsPath });
				await runTests({ 
					extensionDevelopmentPath, extensionTestsPath,
					launchArgs: [
						testWorkspace, 
						"--disable-extensions"
					]
				});
			} catch (err) {
				console.error('Failed to run tests');
				process.exit(1);
			}
		}
	});
}

main();
