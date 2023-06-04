using ConsoleApp1;
using System.Collections.Generic;
using Xunit;

namespace TestProject1 {
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
    }
}
