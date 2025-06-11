using Microsoft.CodeAnalysis.Completion;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading;
using System.Threading.Tasks;
using VBACodeAnalysis;
using VBALanguageServer;
using Xunit;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace TestProject {
	public class TesProf(ITestOutputHelper output) {
		private readonly ITestOutputHelper output = output;
		[Fact]
		public void TestProf1() {
			var code = Helper.getCode("prof_sample1.bas");
			var vbaca = new VBACodeAnalysis.VBACodeAnalysis();
			var preprocVBA = new PreprocVBA();

			var sw = new System.Diagnostics.Stopwatch();
			sw.Start();
			var vbCode = preprocVBA.Rewrite("test", code);
			vbaca.AddDocument("test", vbCode);
			sw.Stop();
			output.WriteLine($"{sw.ElapsedMilliseconds}ms");
		}

		[Fact]
		public void TestProf2() {
			var code = Helper.getCode("prof_sample2.bas");
			var vbaca = new VBACodeAnalysis.VBACodeAnalysis();
			var preprocVBA = new PreprocVBA();

			var sw = new System.Diagnostics.Stopwatch();
			sw.Start();
			var vbCode = preprocVBA.Rewrite("test", code);
			vbaca.AddDocument("test", vbCode);
			sw.Stop();
			output.WriteLine($"{sw.ElapsedMilliseconds}ms");
		}
	}
}
