import * as Mocha from 'mocha';

export function run(): Promise<void> {
	const testFile = process.env["MOCHA_LAUNCH_FILE"]!;

	const mocha = new Mocha({
		ui: 'tdd',
		color: true,
		timeout: 0
	});

	return new Promise((c, e) => {
		// Add files to the test suite
		mocha.addFile(testFile);
		try {
			// Run the mocha test
			mocha.run(failures => {
				if (failures > 0) {
					e(new Error(`${failures} tests failed.`));
				} else {
					c();
				}
			});
		} catch (err) {
			console.error(err);
			e(err);
		}
	});
}
