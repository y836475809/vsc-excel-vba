using ConsoleApp1;
using System.Collections.Generic;
using Xunit;

namespace TestProject1 {
    public class TestDiagOpenClose {
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

            Assert.Equal(1, diagnostics.Count);
            var d0 = diagnostics[0];
            Assert.Contains("Requires 4 arguments", d0.Message);
            Assert.True(d0.Eq(2, 5, 2, 5));
        }

        [Fact]
        public void TestDiagnosticOpenArgOK() {
            var code = MakeCode("Open fname For Output As #1");
            var diagnostics = GetItems(code);

            Assert.Equal(1, diagnostics.Count);
            var d0 = diagnostics[0];
            Assert.Contains("fname", d0.Message);
            Assert.True(d0.Eq(2, 5, 2, 10));
        }

        [Fact]
        public void TestDiagnosticOpenArgFnameOnly() {
            var code = MakeCode("Open fname");
            var diagnostics = GetItems(code);

            Assert.Equal(2, diagnostics.Count);

            var d0 = diagnostics[0];
            var d1 = diagnostics[1];
            Assert.Contains("fname", d0.Message);
            Assert.True(d0.Eq(2, 5, 2, 10));

            Assert.Contains("Requires 4 arguments", d1.Message);
            Assert.True(d1.Eq(2, 11, 2, 11));
        }

        [Fact]
        public void TestDiagnosticOpenArgLess4() {
            var code = MakeCode("Open fname Output #1");
            var diagnostics = GetItems(code);

            Assert.Equal(2, diagnostics.Count);

            var d0 = diagnostics[0];
            var d1 = diagnostics[1];
            Assert.Contains("fname", d0.Message);
            Assert.True(d0.Eq(2, 5, 2, 10));

            Assert.Contains("Requires 4 arguments", d1.Message);
            Assert.True(d1.Eq(2, 0, 2, 4));
        }

        [Fact]
        public void TestDiagnosticOpenArgNoAccess() {
            var code = MakeCode("Open fname For a As #1");
            var diagnostics = GetItems(code);

            Assert.Equal(2, diagnostics.Count);

            var d0 = diagnostics[0];
            var d1 = diagnostics[1];
            Assert.Contains("fname", d0.Message);
            Assert.True(d0.Eq(2, 5, 2, 10));

            Assert.Contains("\"Access\" is required", d1.Message);
            Assert.True(d1.Eq(2, 0, 2, 4));
        }

        [Fact]
        public void TestDiagnosticOpenArgNoFor() {
            var code = MakeCode("Open fname a Output As #1");
            var diagnostics = GetItems(code);

            Assert.Equal(2, diagnostics.Count);

            var d0 = diagnostics[0];
            var d1 = diagnostics[1];
            Assert.Contains("fname", d0.Message);
            Assert.True(d0.Eq(2, 5, 2, 10));

            Assert.Contains("\"For\" is required", d1.Message);
            Assert.True(d1.Eq(2, 0, 2, 4));
        }

        [Fact]
        public void TestDiagnosticOpenArgNoAs() {
            var code = MakeCode("Open fname For Output a #1");
            var diagnostics = GetItems(code);

            Assert.Equal(2, diagnostics.Count);

            var d0 = diagnostics[0];
            var d1 = diagnostics[1];
            Assert.Contains("fname", d0.Message);
            Assert.True(d0.Eq(2, 5, 2, 10));

            Assert.Contains("\"As\" is required", d1.Message);
            Assert.True(d1.Eq(2, 0, 2, 4));
        }

        [Fact]
        public void TestDiagnosticOpenArgNoNum() {
            var code = MakeCode("Open fname For Output As");
            var diagnostics = GetItems(code);

            Assert.Equal(2, diagnostics.Count);

            var d0 = diagnostics[0];
            var d1 = diagnostics[1];
            Assert.Contains("fname", d0.Message);
            Assert.True(d0.Eq(2, 5, 2, 10));

            Assert.Contains("Requires 4 arguments", d1.Message);
            Assert.True(d1.Eq(2, 0, 2, 4));
        }

        [Fact]
        public void TestDiagnosticCloseOK() {
            var code = MakeCode("Close #1");
            var diagnostics = GetItems(code);

            Assert.Empty(diagnostics);
        }

        [Fact]
        public void TestDiagnosticCloseNoArgNo() {
            var code = MakeCode("Close");
            var diagnostics = GetItems(code);

            Assert.Empty(diagnostics);
        }

        [Fact]
        public void TestDiagnosticCloseArgVar() {
            var code = MakeCode("Close fn");
            var diagnostics = GetItems(code);

            Assert.Equal(1, diagnostics.Count);

            var d0 = diagnostics[0];
            Assert.Contains("fn", d0.Message);
            Assert.True(d0.Eq(2, 6, 2, 8));
        }

        [Fact]
        public void TestDiagnosticCloseArgThan1() {
            var code = MakeCode("Close fn #1");
            var diagnostics = GetItems(code);

            Assert.Equal(1, diagnostics.Count);

            var d0 = diagnostics[0];
            Assert.Contains("fn", d0.Message);
            Assert.True(d0.Eq(2, 6, 2, 8));
        }
    }
}
