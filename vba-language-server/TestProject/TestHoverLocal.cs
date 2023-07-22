using System.Collections.Generic;
using VBACodeAnalysis;
using Xunit;

namespace TestProject {
	public class TestHoverLocal {
        private CompletionItem GetItem(string code, int chara) {
            var vbaca = new VBACodeAnalysis.VBACodeAnalysis();
            vbaca.setSetting(new RewriteSetting());
            vbaca.AddDocument("m1", code);
            var srcLine = 11;
            return vbaca.GetHover("m1", srcLine, chara).Result;
        }
        private string MakeCode(string src) {
            var code = @$"Module Module1
private const pri_const_num =10
public const pub_const_num =10
const const_num =10
public const pub_const_str =""""
private pri_num As Long
private pri_non
Dim acc_non As Long
Sub Main()
Dim local_num As Long
Const local_const_num=10
{src}
End Sub
End Module";
            return code;
        }

        [Fact]
        public void TestPrivateConstNum() {
            var code = MakeCode("local_num=pri_const_num+1");
            var item = GetItem(code, "local_num=".Length + 1);
            Assert.Equal(
                "Private Const pri_const_num As Integer = 10",
                item.DisplayText);
        }

        [Fact]
        public void TestPublicConstNum() {
            var code = MakeCode("local_num=pub_const_num+1");
            var item = GetItem(code, "local_num=".Length + 1);
            Assert.Equal(
                "Public Const pub_const_num As Integer = 10",
                item.DisplayText);
        }

        [Fact]
        public void TestNonAccConstNum() {
            var code = MakeCode("local_num=const_num+1");
            var item = GetItem(code, "local_num=".Length + 1);
            Assert.Equal(
                "Private Const const_num As Integer = 10",
                item.DisplayText);
        }

        [Fact]
        public void TestPublicConstStr() {
            var code = MakeCode(@"local_str=pub_const_str & ""a""");
            var item = GetItem(code, "local_str=".Length + 1);
            Assert.Equal(
                @"Public Const pub_const_str As String = """"",
                item.DisplayText);
        }

        [Fact]
        public void TestPrivateNon() {
            var code = MakeCode("local_num=pri_non+1");
            var item = GetItem(code, "local_num=".Length + 1);
            Assert.Equal(
                "Private pri_non As Variant",
                item.DisplayText);
        }

        [Fact]
        public void TestAccNom() {
            var code = MakeCode("local_num=acc_non+1");
            var item = GetItem(code, "ocal_num=".Length + 1);
            Assert.Equal(
                "Private acc_non As Long",
                item.DisplayText);
        }

        [Fact]
        public void TestLocalNum() {
            var code = MakeCode("local_num=pri_num+1");
            var item = GetItem(code, 1);
            Assert.Equal(
                "Local local_num As Long",
                item.DisplayText);
        }

        [Fact]
        public void TestLocalConstNum() {
            var code = MakeCode("local_const_num=pri_num+1");
            var item = GetItem(code, 1);
            Assert.Equal(
                "Local Const local_const_num As Integer = 10",
                item.DisplayText);
        }

    }
}
