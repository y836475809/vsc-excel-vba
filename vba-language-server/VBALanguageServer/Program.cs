using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VBALanguageServer {
	class Program {
		static void Main(string[] args) {
			Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
			Logger.Info(string.Join(" ", args));

			if (!args.Contains("--stdio")) {
				Logger.Info("not found --stdio");
				return;
			}
			var srcDirName = args.Where(x => x.StartsWith("--src_dir_name="))
				.Select(x => x.Replace("--src_dir_name=", ""))
				.FirstOrDefault("");
			if (srcDirName == "") {
				Logger.Info("not found --src_dir_name={dir name}");
				return;
			}

			MainAsync(srcDirName).Wait();
		}

		private static async Task MainAsync(string srcDirName) {
			System.IO.Stream stdin = Console.OpenStandardInput();
			System.IO.Stream stdout = Console.OpenStandardOutput();
			var server = new Server(srcDirName, stdin, stdout);
			await Task.Delay(-1);
		}
	}
}
