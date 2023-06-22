using System.Collections.Generic;
using VBACodeAnalysis;
using Xunit;

namespace TestProject {
	public class TestHoverField {
        private CompletionItem GetItem(string code, string search) {
            var mc = new MyCodeAnalysis();
            mc.setSetting(new RewriteSetting());
            mc.AddDocument("m0", MakeModule());
            mc.AddDocument("c1", MakeClass());
            mc.AddDocument("m1", code);
            var index = code.IndexOf(search);
            return mc.GetHover("m1", "", index + 1).Result;
        }
        private string MakeModule() {
            var code = @$"Public Module m0
Public const pub_const_num =10
Public pub_num As Long
End Module";
            return code;
        }
        private string MakeClass() {
            var code = @$"Public Class c1
Public const pub_const_num =10
Public pub_num As Long
End Class";
            return code;
        }
        private string MakeCode(string src) {
            var code = @$"Module Module1
Sub Main()
Dim local_num As Long
Dim c As New c1
{src}
End Sub
End Module";
            return code;
        }

        [Fact]
        public void TestModuleConstNum() {
            var code = MakeCode("local_num=pub_const_num+1");
            var item = GetItem(code, "pub_const_num+1");
            Assert.Equal(
                "Public Const pub_const_num As Integer = 10",
                item.DisplayText);
        }

        [Fact]
        public void TestModuleNum() {
            var code = MakeCode("local_num=pub_num+1");
            var item = GetItem(code, "pub_num+1");
            Assert.Equal(
                "Public pub_num As Long",
                item.DisplayText);
        }

        [Fact]
        public void TestClassConstNum() {
            var code = MakeCode("local_num=c.pub_const_num+1");
            var item = GetItem(code, "pub_const_num+1");
            Assert.Equal(
                "Public Const pub_const_num As Integer = 10",
                item.DisplayText);
        }

        [Fact]
        public void TestClassNum() {
            var code = MakeCode("local_num=c.pub_num+1");
            var item = GetItem(code, "pub_num+1");
            Assert.Equal(
                "Public pub_num As Long",
                item.DisplayText);
        }
    }
}
