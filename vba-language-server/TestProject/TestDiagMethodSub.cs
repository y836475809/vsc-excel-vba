using VBACodeAnalysis;
using System.Collections.Generic;
using Xunit;

namespace TestProject {
// VBA Sub
// myFunction 正常
// Call myFunction     正常
// myFunction 123 	正常
// myFunction(123)     正常
// Call myFunction 123 	エラー
// Call myFunction(123)    正常

    public class TestDiagMethodSub {
        const int preLine = 3;
        List<string> errorTypes;

        public TestDiagMethodSub() {
            errorTypes = new List<string> { "error" };
        }
        private List<DiagnosticItem> GetDiag(string code) {
            return Helper.GetDiagnostics(MakeSub(code), errorTypes);
        }

        string MakeSub(string stm) {
            return $@"Module Module1
    Sub Main()

        {stm}
    End Sub

    Sub test()
    End Sub
    Sub testArgs(a as Long)
    End Sub
End Module";
        }

    

        [Fact]
		public void TestDiagnosticCallSub1() {
			var items = GetDiag("test");
			Assert.Empty(items);
		}
        [Fact]
        public void TestDiagnosticCallSub2() {
            var items = GetDiag("Call test");
            Assert.Empty(items);
        }
        [Fact]
        public void TestDiagnosticCallSub3() {
            var items = GetDiag("testArgs 123");
            Assert.Empty(items);
        }
        [Fact]
        public void TestDiagnosticCallSub4() {
            var items = GetDiag("Call testArgs(123)");
            Assert.Empty(items);
        }

        [Fact]
        public void TestDiagnosticCallSub5() {
            var items = GetDiag("testArgs(123)");
            Assert.Empty(items);
        }
        [Fact]
        public void TestDiagnosticCallSub6() {
            var items = GetDiag("Call testArgs 123");

            Assert.Equal(2, items.Count);
            {
                var item = items[0];
                Assert.Equal(preLine, item.StartLine);
                Assert.Equal(13, item.StartChara);
                Assert.Equal(preLine, item.EndLine);
                Assert.Equal(21, item.EndChara);
            }
            {
                var item = items[1];
                Assert.Equal(preLine, item.StartLine);
                Assert.Equal(22, item.StartChara);
                Assert.Equal(preLine, item.EndLine);
                Assert.Equal(25, item.EndChara);
            }
        }
    }
}
