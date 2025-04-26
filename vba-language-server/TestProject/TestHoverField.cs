using System.Collections.Generic;
using System.Linq;
using VBACodeAnalysis;
using Xunit;

namespace TestProject {
	public class TestHoverField {
        private VBAHover GetItem(string code, int chara) {
            var vbaca = new VBACodeAnalysis.VBACodeAnalysis();
            vbaca.setSetting(new RewriteSetting());
            vbaca.AddDocument("m0", MakeModule());
            vbaca.AddDocument("c1", MakeClass());
            vbaca.AddDocument("m1", code);
            var srcLine = 4;
            return vbaca.GetHover("m1", srcLine, chara).Result;
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
            var hover = GetItem(code, "local_num=".Length + 1);
			Assert.Equal(3, hover.Contents.Count);
			var act = hover.Contents.Select(x => x.Value);
			Assert.Equal(
				["Public Const pub_const_num As Integer = 10", "@return Integer", "@kind Field"],
				[.. act]
			 );
		}

        [Fact]
        public void TestModuleNum() {
            var code = MakeCode("local_num=pub_num+1");
            var hover = GetItem(code, "local_num=".Length + 1);
			Assert.Equal(3, hover.Contents.Count);
			var act = hover.Contents.Select(x => x.Value);
			Assert.Equal(
				["Public pub_num As Long", "@return Long", "@kind Field"],
				[.. act]
			 );
		}

        [Fact]
        public void TestClassConstNum() {
            var code = MakeCode("local_num=c.pub_const_num+1");
            var hover = GetItem(code, "local_num=c.".Length + 1);
			Assert.Equal(3, hover.Contents.Count);
			var act = hover.Contents.Select(x => x.Value);
			Assert.Equal(
				["Public Const pub_const_num As Integer = 10", "@return Integer", "@kind Field"],
				[.. act]
			 );
        }

        [Fact]
        public void TestClassNum() {
            var code = MakeCode("local_num=c.pub_num+1");
            var hover = GetItem(code, "local_num=c.".Length + 1);
			Assert.Equal(3, hover.Contents.Count);
			var act = hover.Contents.Select(x => x.Value);
            Assert.Equal(
                ["Public pub_num As Long", "@return Long", "@kind Field"],
				[..act]
			 );
		}
    }
}
