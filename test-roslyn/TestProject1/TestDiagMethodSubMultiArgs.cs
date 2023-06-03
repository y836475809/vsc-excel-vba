using ConsoleApp1;
using System.Collections.Generic;
using Xunit;

namespace TestProject1 {
    // VBA Sub
    // testArgs 1, 2       '正常
    // testArgs(1, 2)      'エラー
    // Call testArgs 1, 2  'エラー
    // Call testArgs(1, 2) '正常

    public class TestDiagMethodSubMultiArgs {
        const int preLine = 3;
        List<string> errorTypes;

        public TestDiagMethodSubMultiArgs() {
            errorTypes = new List<string> { "error" };
        }
        private List<DiagnosticItem> GetDiag(string code) {
            return Helper.GetDiagnostics(MakeSub(code), errorTypes);
        }

        string MakeSub(string stm) {
            return $@"Module Module1
    Sub Main()
        Dim ret As Long
        {stm}
    End Sub

    Sub testArgs(a as Long, b As Long)
    End Sub
End Module";
        }

        [Fact]
        public void TestDiagnosticCallSub3() {
            var items = GetDiag("testArgs 1,2");
            Assert.Empty(items);
        }
        [Fact]
        public void TestDiagnosticCallSub4() {
            var items = GetDiag("testArgs(1,2)");
            Assert.Single(items);
        }

        [Fact]
        public void TestDiagnosticCallSub5() {
            var items = GetDiag("Call testArgs 1,2");
            Assert.Equal(3, items.Count);
        }
        [Fact]
        public void TestDiagnosticCallSub6() {
            var items = GetDiag("Call testArgs(1, 2)");
            Assert.Empty(items);
        }
    }
}
