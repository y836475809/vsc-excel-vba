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
using static System.Runtime.CompilerServices.RuntimeHelpers;

namespace TestProject {
	public class TesRewriteProperty(ITestOutputHelper output) {
		private readonly ITestOutputHelper output = output;

		[Fact]
		public void TestRewrite() {
			var code = Helper.getCode("test_property.bas");
			var preprocVBA = new PreprocVBA();
			var actCode = preprocVBA.Rewrite("test", code);
			var preCode = @"
Public Property     Name1() As String
Set : End Set
Get
Name1 = LCase(Name)
End Get:End Property

Private Sub            Name1(argName As String)
Dim a As String
a = argName
Sub Function

Private Sub set_Name3(argName As String)
Me.Name = argName
End Sub
Public Property Name3 As String

Property ReadOnly Name2() As String
Get
Name2 = LCase(Name)
End Get:End Property

";
			Helper.AssertCode(preCode, actCode);
		}
	}
}
