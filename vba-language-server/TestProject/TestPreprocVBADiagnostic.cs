using Microsoft.CodeAnalysis.Operations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text;
using System.Threading.Tasks;
using TestProject;
using VBACodeAnalysis;
using Xunit;

namespace TestPreprocVBA {
	public class TestPreprocVBADiagnostic {
		[Theory]
		[InlineData("Open fp For Output As #1", 4, "Open")]
		[InlineData("Open fp For Output Access Read As #1", 4, "Open")]
		[InlineData("Open fp For Output Access Read Shared As #1", 4, "Open")]
		[InlineData("Open fp For Output Access Read Write Lock Read As #1", 4, "Open")]
		[InlineData("Open fp For Output Access Read Lock Read As #1", 4, "Open")]
		[InlineData("Open fp For Output Access Read Write Shared As #1", 4, "Open")]
		[InlineData("Close #1", 5, "Close")]
		[InlineData("Print #1, \"test\"", 5, "Print")]
		[InlineData("Print #1,", 5, "Print")]
		[InlineData("Print #1, \"print 1\"; Tab ; \"print 2\"", 5, "Print")]
		[InlineData("Print #1, Spc(5) ; \"5spaces\"", 5, "Print")]
		[InlineData("Print #1, Tab(10, 1, Spc(1, 3)) ; \"10tab\"", 5, "Print")]
		[InlineData("Write #1, \"test\"", 5, "Write")]
		[InlineData("Write #1,", 5, "Write")]
		[InlineData("Write #1, \"print 1\"; Tab ; \"print 2\"", 5, "Write")]
		[InlineData("Write #1, Spc(5) ; \"5spaces\"", 5, "Write")]
		[InlineData("Write #1, Tab(10, 1, Spc(1, 3)) ; \"10tab\"", 5, "Write")]
		[InlineData("Input #1, v1", 5, "Input")]
		[InlineData("Input #1, v1, v2", 5, "Input")]
		[InlineData("Line Input #1, v1", 10, "Line_Input")]
		public void TestIgnoreDiagnostics(string code, int EndCol, string text) {
			var pp = new TestPreprocVBA();
			pp.Rewrite("test", code);
			var ignores = pp.GetIgnoreDiagnostics("test");
			Assert.Single(ignores);
			var act = ignores[0];
			Assert.Equal(0, act.Start.Item1);
			Assert.Equal(0, act.Start.Item2);
			Assert.Equal(0, act.End.Item1);
			Assert.Equal(EndCol, act.End.Item2);
			Assert.Equal(text, act.Code);
		}

		public static IEnumerable<object[]> InnutOutputParams =>
		[
				[ @"
Print #1 text", new VBADiagnostic[]{
					new (){
						ID = "VBA_print",
						Severity = "Error",
						Message = "",
						Start = (1, 0), End = (1, 5)
					}
				} ],
				[ @"
Print fn, text", new VBADiagnostic[]{
					new (){
						ID = "BC30451",
						Severity = "Error",
						Message = "",
						Start = (1, 6), End = (1, 8)
					}
				} ],
				[ @"
Print fn, text2", new VBADiagnostic[]{
					new (){
							ID = "BC30451",
							Severity = "Error",
							Message = "",
							Start = (1, 6), End = (1, 8)
						},
						new (){
							ID = "BC30451",
							Severity = "Error",
							Message = "",
							Start = (1, 10), End = (1, 15)
						}
				} ],
				[ @"
Write #1 text", new VBADiagnostic[]{
					new (){
							ID = "VBA_write",
							Severity = "Error",
							Message = "",
							Start = (1, 0), End = (1, 5)
						}
				} ],
				[ @"
Write fn, text", new VBADiagnostic[]{
					new (){
					ID = "BC30451",
					Severity = "Error",
					Message = "",
					Start = (1, 6), End = (1, 8)
					}
				}],
				[ @"
Write fn, text2", new VBADiagnostic[]{
					new (){
							ID = "BC30451",
							Severity = "Error",
							Message = "",
							Start = (1, 6), End = (1, 8)
						},
						new (){
							ID = "BC30451",
							Severity = "Error",
							Message = "",
							Start = (1, 10), End = (1, 15)
						}
				} ],
				[ @"
Close, #1", new VBADiagnostic[]{
					new (){
							ID = "VBA_close",
							Severity = "Error",
							Message = "",
							Start = (1, 0), End = (1, 5)
						}
				} ],
				[ @"
Close fn", new VBADiagnostic[]{
					new (){
							ID = "BC30451",
							Severity = "Error",
							Message = "",
							Start = (1, 6), End = (1, 8)
						}
				} ],
				[ @"
Input #1,", new VBADiagnostic[]{
					new (){
							ID = "VBA_input",
							Severity = "Error",
							Message = "",
							Start = (1, 0), End = (1, 5)
						},
						new  (){
							ID = "BC30201",
							Severity = "Error",
							Message = "",
							Start = (1, 9), End = (1, 9)
						}
				} ],
				[ @"
Input #1 text", new VBADiagnostic[]{
					new (){
							ID = "VBA_input",
							Severity = "Error",
							Message = "",
							Start = (1, 0), End = (1, 5)
						}
				} ],
				[ @"
Input fn, text, text", new VBADiagnostic[]{
					new (){
							ID = "BC30451",
							Severity = "Error",
							Message = "",
							Start = (1, 6), End = (1, 8)
						}
				} ],
				[ @"
Input fn, text, text2", new VBADiagnostic[]{
					new (){
							ID = "BC30451",
							Severity = "Error",
							Message = "",
							Start = (1, 6), End = (1, 8)
						},
						new (){
							ID = "BC30451",
							Severity = "Error",
							Message = "",
							Start = (1, 16), End = (1, 21)
						}
				} ],

				[ @"'
Line Input #1,", new VBADiagnostic[]{
					new () {
							ID = "VBA_line_input",
							Severity = "Error",
							Message = "",
							Start = (1, 0), End = (1, 10)
						},
						new () {
							ID = "BC30201",
							Severity = "Error",
							Message = "",
							Start = (1, 14), End = (1, 14)
						}
				} ],
				[ @"'
Line    Input #1,", new VBADiagnostic[]{
					new (){
							ID = "VBA_line_input",
							Severity = "Error",
							Message = "",
							Start = (1, 0), End = (1, 10)
						},
						new (){
							ID = "BC30201",
							Severity = "Error",
							Message = "",
							Start = (1, 17), End = (1, 17)
						}
				} ],
				[ @"
Line Input #1 text", new VBADiagnostic[]{
					new (){
							ID = "VBA_line_input",
							Severity = "Error",
							Message = "",
							Start = (1, 0), End = (1, 10)
						}
				} ],
				[ @"
Line Input #1, text, text", new VBADiagnostic[]{
					new (){
							ID = "VBA_line_input",
							Severity = "Error",
							Message = "",
							Start = (1, 0), End = (1, 10)
						}
				} ],
				[ @"
Line Input fn, text", new VBADiagnostic[]{
					new (){
							ID = "BC30451",
							Severity = "Error",
							Message = "",
							Start = (1, 11), End = (1, 13)
						}
				} ],
				[ @"
Line Input fn, text2", new VBADiagnostic[]{
					new (){
							ID = "BC30451",
							Severity = "Error",
							Message = "",
							Start = (1, 11), End = (1, 13)
						},
						new (){
							ID = "BC30451",
							Severity = "Error",
							Message = "",
							Start = (1, 15), End = (1, 20)
						}
				}],
		];
		[Theory]
		[MemberData(nameof(InnutOutputParams))]
		public async Task TestDiagnosticInputOutputError(string output, VBADiagnostic[] pre) {
			var code = $@"Attribute VB_Name = ""test""
Sub Main()
Dim text As String:text=""msg""
{output}
End Sub";
			foreach (var item in pre) {
				item.Start = (item.Start.Item1 + 3, item.Start.Item2);
				item.End = (item.End.Item1 + 3, item.End.Item2);
			}
			var vbaca = new VBACodeAnalysis.VBACodeAnalysis();
			vbaca.setSetting(new RewriteSetting());
			var vbCode = vbaca.Rewrite("test", code);
			vbaca.AddDocument("test", vbCode);
			var diagnosticParams = await vbaca.GetDiagnostics("test");
			Helper.AssertDiagnostics([..pre], diagnosticParams);
		}

		public static IEnumerable<object[]> OpenParams =>
		[
				[ @"
Open fp For Output", new VBADiagnostic[]{
						new (){
							ID = "VBA_open",
							Severity = "Error",
							Message = "",
							Start = (1, 0), End = (1, 4)
						}
				} ],
				[ @"
Open fp For Output As fn", new VBADiagnostic[]{
					new (){
							ID = "BC30451",
							Severity = "Error",
							Message = "",
							Start = (1, 22), End = (1, 24)
						}
				} ],
				[ @"
Open f1 For Output As fn", new VBADiagnostic[]{
						new (){
							ID = "BC30451",
							Severity = "Error",
							Message = "",
							Start = (1, 5), End = (1, 7)
						},
						new (){
							ID = "BC30451",
							Severity = "Error",
							Message = "",
							Start = (1, 22), End = (1, 24)
						}
				} ],
				[ @"
Open fp As #1", new VBADiagnostic[]{
					new (){
							ID = "VBA_open",
							Severity = "Error",
							Message = "",
							Start = (1, 0), End = (1, 4)
						}
				} ],
				[ @"
Open fp For xOutput As #1",  new VBADiagnostic[]{
					new (){
							ID = "VBA_open",
							Severity = "Error",
							Message = "",
							Start = (1, 0), End = (1, 4)
						}
				} ],
				[ @"
Open fp For Output Access As #1", new VBADiagnostic[]{
					new (){
							ID = "VBA_open",
							Severity = "Error",
							Message = "",
							Start = (1, 0), End = (1, 4)
						}
				} ],
				[ @"
Open fp For Output Access xRead As #1", new VBADiagnostic[]{
					new (){
							ID = "VBA_open",
							Severity = "Error",
							Message = "",
							Start = (1, 0), End = (1, 4)
						}
				} ],
				[ @"
Open fp For Output Access Read xWrite As #1", new VBADiagnostic[]{
					new (){
							ID = "VBA_open",
							Severity = "Error",
							Message = "",
							Start = (1, 0), End = (1, 4)
						}
				} ],
				[ @"
Open fp For Output xShared As #1", new VBADiagnostic[]{
					new (){
							ID = "VBA_open",
							Severity = "Error",
							Message = "",
							Start =(1, 0), End =(1, 4)
						}
				} ],
				[ @"
Open fp For Output Lock xRead As #1", new VBADiagnostic[]{
					new (){
							ID = "VBA_open",
							Severity = "Error",
							Message = "",
							Start =(1, 0), End =(1, 4)
						}
				} ],
				[ @"
Open fp For Output Access Read Write xShared As #1", new VBADiagnostic[]{
					new (){
							ID = "VBA_open",
							Severity = "Error",
							Message = "",
							Start =(1, 0), End =(1, 4)
						}
				}],
				[ @"
Open fp For Output Access Read Write Lock xRead As #1", new VBADiagnostic[]{
					new (){
							ID = "VBA_open",
							Severity = "Error",
							Message = "",
							Start =(1, 0), End =(1, 4)
						}
					}
				],
				[ @"
Open f1 For xOutput Access xRead Write Lock xRead As fn", new VBADiagnostic[]{
					new (){
							ID = "VBA_open",
							Severity = "Error",
							Message = "",
							Start =(1, 0), End =(1, 4)
						}
					}
				],
				[ @"
Open f1 For xOutput Access Read Write Lock Read As fn", new VBADiagnostic[]{
					new () {
							ID = "VBA_open",
							Severity = "Error",
							Message = "",
							Start = (1, 0), End = (1, 4)
						},
						new (){
							ID = "BC30451",
							Severity = "Error",
							Message = "",
							Start = (1, 5), End = (1, 7)
						},
						new (){
							ID = "BC30451",
							Severity = "Error",
							Message = "",
							Start =(1, 51), End =(1, 53)
						}
				}
			]
		];
		[Theory]
		[MemberData(nameof(OpenParams))]
		public async Task TestDiagnosticOpenError(string output, VBADiagnostic[] pre) {
			var module = @"Attribute VB_Name = ""module1""
Public fp As String
";
			var code = $@"Attribute VB_Name = ""test""
Sub Main()
Dim text As String:text=""msg""
{output}
End Sub";
			foreach (var item in pre) {
				item.Start = (item.Start.Item1 + 3, item.Start.Item2);
				item.End = (item.End.Item1 + 3, item.End.Item2);
			}
			var vbaca = new VBACodeAnalysis.VBACodeAnalysis();
			vbaca.setSetting(new RewriteSetting());
			var vbModule = vbaca.Rewrite("m1", module);
			var vbCode = vbaca.Rewrite("test", code);
			vbaca.AddDocument("m1", vbModule);
			vbaca.AddDocument("test", vbCode);
			var diagnosticParams = await vbaca.GetDiagnostics("test");
			Helper.AssertDiagnostics([.. pre], diagnosticParams);
		}
	}
}
