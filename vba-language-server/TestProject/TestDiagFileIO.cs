using VBACodeAnalysis;
using System.Collections.Generic;
using Xunit;

namespace TestProject {
    public class TestDiagFileIO {
        private List<DiagnosticItem> GetItems(string code) {
            var mc = new MyCodeAnalysis();
            mc.setSetting(new RewriteSetting());
            mc.AddDocument("m1", code);
            return mc.GetDiagnostics("m1").Result;
        }

        private string MakeCode(string src) {
            var code = @$"Module Module1
Sub Main()
{src}
End Sub
End Module";
            return code;
        }

        [Fact]
        public void TestDiagnosticOpenNoArgs() {
            var code = MakeCode("Open");
            var diagnostics = GetItems(code);

            Assert.Single(diagnostics);
            Assert.Contains("Open", diagnostics[0].Message);
        }

        [Fact]
        public void TestDiagnosticOpenArgFlieNum() {
            var code = MakeCode("Open fname For Output As #1");
            var diagnostics = GetItems(code);

            Assert.Single(diagnostics);
            Assert.Contains("fname", diagnostics[0].Message);
        }

        [Fact]
        public void TestDiagnosticOpenArgVal() {
            var code = MakeCode("Open fname For Output As fnum");
            var diagnostics = GetItems(code);

            Assert.Single(diagnostics);
            Assert.Contains("fname", diagnostics[0].Message);
        }

        [Fact]
        public void TestDiagnosticOpenNewline() {
            var code = MakeCode(@"Open fname & ""_post"" _
                                                For Output As fnum");
            var diagnostics = GetItems(code);

            Assert.Single(diagnostics);
            Assert.Contains("fname", diagnostics[0].Message);
        }

        [Fact]
        public void TestDiagnosticCloseFileNum() {
            var code = MakeCode("Close #1");
            var diagnostics = GetItems(code);

            Assert.Single(diagnostics);
            Assert.Contains("Close", diagnostics[0].Message);
        }

        [Fact]
        public void TestDiagnosticCloseNoArgNo() {
            var code = MakeCode("Close");
            var diagnostics = GetItems(code);

            Assert.Single(diagnostics);
            Assert.Contains("Close", diagnostics[0].Message);
        }

        [Fact]
        public void TestDiagnosticCloseArgVar() {
            var code = MakeCode("Close fn");
            var diagnostics = GetItems(code);

            Assert.Equal(2, diagnostics.Count);
            Assert.Contains("Close", diagnostics[0].Message);
            Assert.Contains("fn", diagnostics[1].Message);
        }

        [Fact]
        public void TestDiagnosticPrintFnum() {
            var code = MakeCode(@"Print #1, ""aaa""");
            var diagnostics = GetItems(code);

            Assert.Single(diagnostics);
            Assert.Contains("Print", diagnostics[0].Message);
        }

        [Fact]
        public void TestDiagnosticPrintFnumVal() {
            var code = MakeCode(@"Print fnum, ""aaa""");
            var diagnostics = GetItems(code);

            Assert.Equal(2, diagnostics.Count);
            Assert.Contains("Print", diagnostics[0].Message);
            Assert.Contains("fnum", diagnostics[1].Message);
        }

        [Fact]
        public void TestDiagnosticPrintTextVal() {
            var code = MakeCode(@"Print #1, text");
            var diagnostics = GetItems(code);

            Assert.Equal(2, diagnostics.Count);
            Assert.Contains("Print", diagnostics[0].Message);
            Assert.Contains("text", diagnostics[1].Message);
        }

        [Fact]
        public void TestDiagnosticPrintFnumDecVal() {
            var code = MakeCode(@"Dim fnum As Long
Print fnum, ""aaa""");
            var diagnostics = GetItems(code);

            Assert.Single(diagnostics);
            Assert.Contains("Print", diagnostics[0].Message);
        }
    }
}
