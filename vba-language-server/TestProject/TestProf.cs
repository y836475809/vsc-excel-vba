using Microsoft.CodeAnalysis.Completion;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading;
using System.Threading.Tasks;
using VBACodeAnalysis;
using VBALanguageServer;
using VBARewrite;
using Xunit;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace TestProject {
	public class TestProf(ITestOutputHelper output) {
		private readonly ITestOutputHelper output = output;

		[Fact]
		public void TestProfRewriteAndAddDocument() {
			var dirPath = GetDirPath();
			var filePaths = Directory.GetFiles(dirPath, "*.bas");
			foreach (var filePath in filePaths) {
				var code = GetCode(filePath);
				var vbaca = new VBACodeAnalysis.VBACodeAnalysis();
				var rewriter = new VBARewriter();

				var sw = new System.Diagnostics.Stopwatch();
				sw.Start();
				var vbCode = rewriter.Rewrite("test", code);
				vbaca.AddDocument("test", vbCode);
				sw.Stop();
				var fileName = Path.GetFileName(filePath);
				output.WriteLine($"{fileName} : {sw.ElapsedMilliseconds}ms");
			}
		}

		private static string GetCode(string filePath) {
			Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
			return Util.GetCode(filePath);
		}

		private static string GetDirPath([CallerFilePath] string filePath = "") {
			return Path.Combine(Path.GetDirectoryName(filePath), "prof_code");
		}
	}
}
