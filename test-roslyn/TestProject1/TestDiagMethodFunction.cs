﻿using ConsoleApp1;
using System.Collections.Generic;
using Xunit;

namespace TestProject1 {
// VBA Function
// myFunction 正常
// Call myFunction     正常
// myFunction 123 	正常
// myFunction(123)     正常
// Call myFunction 123 	エラー
// Call myFunction(123)    正常
// ret = myFunction    正常
// ret = Call myFunction エラー
// ret = myFunction 123 	エラー
// ret = myFunction(123)   正常
// ret = Call myFunction 123 	エラー
// ret = Call myFunction(123)  エラー

    public class TestDiagMethodFunction {
        const int preLine = 3;
        List<string> errorTypes;
        public TestDiagMethodFunction() {
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

    Function test() As Long
        test = 10
    End Function
    Function testArgs(a as Long) As Long
        testArgs = 10
    End Function
End Module";
        }

        [Fact]
        public void TestDiagnosticCallFunc1() {
            var items = GetDiag("test");
            Assert.Empty(items);
        }
        [Fact]
        public void TestDiagnosticCallFunc2() {
            var items = GetDiag("Call test");
            Assert.Empty(items);
        }
        [Fact]
        public void TestDiagnosticCallFunc3() {
            var items = GetDiag("testArgs 123");
            Assert.Empty(items);
        }
        [Fact]
        public void TestDiagnosticCallFunc4() {
            var items = GetDiag("Call testArgs(123)");
            Assert.Empty(items);
        }
        [Fact]
        public void TestDiagnosticCallFunc5() {
            var items = GetDiag("testArgs(123)");
            Assert.Empty(items);
        }
        [Fact]
        public void TestDiagnosticCallFunc6() {
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

        [Fact]
        public void TestDiagnosticCallFuncRet1() {
            var items = GetDiag("ret=test");
            Assert.Empty(items);
        }
        [Fact]
        public void TestDiagnosticCallFuncRet2() {
            var items = GetDiag("ret=Call test");
            
            Assert.Single(items);
            var item = items[0];
            Assert.Equal(preLine, item.StartLine);
            Assert.Equal(12, item.StartChara);
            Assert.Equal(preLine, item.EndLine);
            Assert.Equal(12, item.EndChara);
        }
        [Fact]
        public void TestDiagnosticCallFuncRet3() {
            var items = GetDiag("ret=testArgs 123");
            
            Assert.Equal(2, items.Count);
            {
                var item = items[0];
                Assert.Equal(preLine, item.StartLine);
                Assert.Equal(12, item.StartChara);
                Assert.Equal(preLine, item.EndLine);
                Assert.Equal(20, item.EndChara);
            }
            {
                var item = items[1];
                Assert.Equal(preLine, item.StartLine);
                Assert.Equal(21, item.StartChara);
                Assert.Equal(preLine, item.EndLine);
                Assert.Equal(24, item.EndChara);
            }
        }
        [Fact]
        public void TestDiagnosticCallFuncRet4() {
            var items = GetDiag("ret=testArgs(123)");
            Assert.Empty(items);
        }
        [Fact]
        public void TestDiagnosticCallFuncRet5() {
            var items = GetDiag("ret=Call testArgs 123");
            
            Assert.Single(items);
            var item = items[0];
            Assert.Equal(preLine, item.StartLine);
            Assert.Equal(12, item.StartChara);
            Assert.Equal(preLine, item.EndLine);
            Assert.Equal(12, item.EndChara);
        }
        [Fact]
        public void TestDiagnosticCallFuncRet6() {
            var items = GetDiag("ret=Call testArgs(123)");
            
            Assert.Single(items);
            var item = items[0];
            Assert.Equal(preLine, item.StartLine);
            Assert.Equal(12, item.StartChara);
            Assert.Equal(preLine, item.EndLine);
            Assert.Equal(12, item.EndChara);
        }
    }
}
