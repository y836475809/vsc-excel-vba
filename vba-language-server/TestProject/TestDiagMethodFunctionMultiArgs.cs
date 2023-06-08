using VBACodeAnalysis;
using System.Collections.Generic;
using Xunit;

namespace TestProject {
    // VBA Function
    // testArgs 1, 2        '正常
    // testArgs(1,2)       'エラー
    // Call testArgs 1,2    'エラー
    // Call testArgs(1, 2)  '正常

    // Dim ret As Long
    // ret = testArgs 1,2       'エラー
    // ret = testArgs(1, 2)     '正常
    // ret = Call testArgs 1,2  'エラー
    // ret = Call testArgs(1,2) 'エラー

    public class TestDiagMethodFunctionMUltiArgs {
        List<string> errorTypes;
        public TestDiagMethodFunctionMUltiArgs() {
            errorTypes = new List<string> { "error" };
        }
        private List<DiagnosticItem> GetDiag(string code) {
            return Helper.GetDiagnostics(MakeFunc(code), errorTypes);
        }

        string MakeFunc(string stm) {
                return $@"Module Module1
    Sub Main()
        Dim ret As Long
        {stm}
    End Sub
    Function testArgs(a As Long, b As Long) As Long
        testArgs = 10
    End Function
    Sub Add(a as Long)
    End Sub
    Sub Eq(a as Boolean)
    End Sub
End Module";
        }

        [Fact]
        public void TestDiagnosticCallFunc1() {
            var items = GetDiag("testArgs 1, 2");
            Assert.Empty(items);
        }
        [Fact]
        public void TestDiagnosticCallFunc2() {
            var items = GetDiag("testArgs(1,2)");
            Assert.Single(items);
        }
        [Fact]
        public void TestDiagnosticCallFunc3() {
            var items = GetDiag("Call testArgs 1,2");
            Assert.Equal(3, items.Count);
        }
        [Fact]
        public void TestDiagnosticCallFunc4() {
            var items = GetDiag("Call testArgs(1, 2)");
            Assert.Empty(items);
        }
        [Fact]
        public void TestDiagnosticCallFunc5() {
            var items = GetDiag("ret = testArgs 1,2");
            Assert.Equal(3, items.Count);
        }
        [Fact]
        public void TestDiagnosticCallFunc6() {
            var items = GetDiag("ret = testArgs(1, 2)");
            Assert.Empty(items);
        }

        [Fact]
        public void TestDiagnosticCallFuncRet1() {
            var items = GetDiag(" ret = Call testArgs 1,2");
            Assert.Single(items);
        }
        [Fact]
        public void TestDiagnosticCallFuncRet2() {
            var items = GetDiag("ret = Call testArgs(1,2)");
            Assert.Single(items);
        }

        [Fact]
        public void TestDiagnosticCallFuncIf() {
            var code = @"If testArgs(1,2) Then
End If";
            var items = GetDiag(code);
            Assert.Empty(items);
        }
        [Fact]
        public void TestDiagnosticCallFuncNEst() {
            var code = @"Call testArgs(1,testArgs(1,2))";
            var items = GetDiag(code);
            Assert.Empty(items);
        }

        [Fact]
        public void TestDiagnosticCallFuncAdd() {
            var code = @"Add testArgs(1,2)";
            var items = GetDiag(code);
            Assert.Empty(items);
        }

        [Fact]
        public void TestDiagnosticCallFuncAddError() {
            var code = @"Add testArgs 1,2 ";
            var items = GetDiag(code);
            Assert.Equal(4, items.Count);
        }

        [Fact]
        public void TestDiagnosticCallFuncEq() {
            var code = @"Eq testArgs(1,2) = 1";
            var items = GetDiag(code);
            Assert.Empty(items);
        }

        [Fact]
        public void TestDiagnosticCallFuncEqVar() {
            var code = @"Dim eqret As Long
eqret = 1
Eq testArgs(1,2) = eqret";
            var items = GetDiag(code);
            Assert.Empty(items);
        }
    }
}
