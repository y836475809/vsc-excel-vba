using System;
using System.Collections.Generic;
using System.Linq;
using VBACodeAnalysis;
using Xunit;

namespace TestProject {
	public class TestHoverLocal {
        private VBAHover GetItem(string code, int chara) {
            var vbaca = new VBACodeAnalysis.VBACodeAnalysis();
            vbaca.AddDocument("m1", code);
            var srcLine = 11;
            return vbaca.GetHover("m1", srcLine, chara).Result;
        }

		private VBAHover GetItem(string code, int line, int chara) {
			var vbaca = new VBACodeAnalysis.VBACodeAnalysis();
			vbaca.AddDocument("m1", code);
			return vbaca.GetHover("m1", line, chara).Result;
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
            var hover = GetItem(code, "local_num=".Length + 1);
			var act = hover.Contents.Select(x => x.Value);
			Assert.Equal(
				["Private Const pri_const_num As Integer = 10", "@kind Field"],
				[.. act]
			 );
		}

        [Fact]
        public void TestPublicConstNum() {
            var code = MakeCode("local_num=pub_const_num+1");
            var hover = GetItem(code, "local_num=".Length + 1);
			var act = hover.Contents.Select(x => x.Value);
			Assert.Equal(
				["Public Const pub_const_num As Integer = 10", "@kind Field"],
				[.. act]
			 );
		}

        [Fact]
        public void TestNonAccConstNum() {
            var code = MakeCode("local_num=const_num+1");
            var hover = GetItem(code, "local_num=".Length + 1);
			var act = hover.Contents.Select(x => x.Value);
			Assert.Equal(
				["Private Const const_num As Integer = 10", "@kind Field"],
				[.. act]
			 );
		}

        [Fact]
        public void TestPublicConstStr() {
            var code = MakeCode(@"local_str=pub_const_str & ""a""");
            var hover = GetItem(code, "local_str=".Length + 1);
			var act = hover.Contents.Select(x => x.Value);
			Assert.Equal(
				["Public Const pub_const_str As String = \"\"", "@kind Field"],
				[.. act]
			 );
		}

        [Fact]
        public void TestPrivateNon() {
            var code = MakeCode("local_num=pri_non+1");
            var hover = GetItem(code, "local_num=".Length + 1);
			var act = hover.Contents.Select(x => x.Value);
			Assert.Equal(
				["Private pri_non As Variant", "@kind Field"],
				[.. act]
			 );
        }

        [Fact]
        public void TestAccNom() {
            var code = MakeCode("local_num=acc_non+1");
            var hover = GetItem(code, "ocal_num=".Length + 1);
			var act = hover.Contents.Select(x => x.Value);
			Assert.Equal(
				["Private acc_non As Long", "@kind Field"],
				[.. act]
			 );
		}

        [Fact]
        public void TestLocalNum() {
            var code = MakeCode("local_num=pri_num+1");
            var hover = GetItem(code, 1);
			var act = hover.Contents.Select(x => x.Value);
			Assert.Equal(
				["Local local_num As Long", "@kind Local"],
				[.. act]
			 );
		}

        [Fact]
        public void TestLocalConstNum() {
            var code = MakeCode("local_const_num=pri_num+1");
            var hover = GetItem(code, 1);
			var act = hover.Contents.Select(x => x.Value);
			Assert.Equal(
				["Local Const local_const_num As Integer = 10", "@kind Local"],
				[.. act]
			 );
		}

		[Theory]
		[InlineData("Dim ary() As Long",			  "Local ary As Long()")]
		[InlineData("Dim ary(1, 2) As Long",	  "Local ary As Long(,)")]
		[InlineData("Dim ary(1, 2, 3) As Long", "Local ary As Long(,,)")]
		[InlineData("Dim ary(1 To 2) As Long",  "Local ary As Long()")]
		[InlineData("Dim ary(1 To 2, 1 To 2) As Long", "Local ary As Long(,)")]
		[InlineData("Dim ary(1 To 2, 1 To 2, 1 To 2) As Long", "Local ary As Long(,,)")]
		public void TestLocalDimArray(string text, string expContent) {
			var codes = new string[] {
$@"Module Module1
Sub Main()
{text}
ary
End Sub
End Module",
$@"Class class1
Sub Main()
{text}
ary
End Sub
End Class"};
			foreach (var code in codes) {
				var hover = GetItem(code, 3, 1);
				var act1 = hover.Contents.Select(x => x.Value);
				Assert.Equal(
					[expContent, "@kind Local"],
					[.. act1]
				 );
			}
		}

		[Theory]
		[InlineData("Dim ary() As Long:  ReDim ary(2)", "Local ary As Long()")]
		[InlineData("Dim ary(,) As Long: ReDim ary(1, 2)", "Local ary As Long(,)")]
		[InlineData("Dim ary(,,) As Long:ReDim ary(1, 2, 3)", "Local ary As Long(,,)")]
		[InlineData("Dim ary() As Long:  ReDim ary(1 To 2)", "Local ary As Long()")]
		[InlineData("Dim ary(,) As Long: ReDim ary(1 To 2, 1 To 2)", "Local ary As Long(,)")]
		[InlineData("Dim ary(,,) As Long:ReDim ary(1 To 2, 1 To 2, 1 To 2)", "Local ary As Long(,,)")]
		public void TestLocalReDimArray(string text, string expContent) {
			var codes = new string[] {
$@"Module Module1
Sub Main()
{text}
ary
End Sub
End Module",
$@"Class class1
Sub Main()
{text}
ary
End Sub
End Class"};
			foreach (var code in codes) {
				var hover = GetItem(code, 3, 1);
				var act1 = hover.Contents.Select(x => x.Value);
				Assert.Equal(
					[expContent, "@kind Local"],
					[.. act1]
				 );
			}
		}
	}
}
