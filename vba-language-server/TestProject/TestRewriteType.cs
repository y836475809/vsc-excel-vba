using VBACodeAnalysis;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace TestProject {
	public class TestRewriteType {

        private string GetCode(string acst, string st, string acx, string acy) {
            return @$"Public Module Module1
  {acst} {st} Point
{acx} X As Long
{acy} Y As Long
    End {st}
X As Long
End Module";
        }

        //string rewriteType(string code) {
        //    var doc = Helper.MakeDoc(code);
        //    var rew = new Rewrite(new RewriteSetting());
        //    var root = rew.TypeStatement(doc.GetSyntaxRootAsync().Result);
        //    return root.ToString();
        //}

        (string, Dictionary<int, (int, int)>) rewriteType(string code) {
            var doc = Helper.MakeDoc(code);
            var rew = new Rewrite(new RewriteSetting());
            var root = rew.TypeStatement(doc.GetSyntaxRootAsync().Result);
            return (root.ToString(), rew.charaOffsetDict);
        }

        [Fact]
        public void TestType() {
            var code = GetCode("", "Type", "", "");
            (var act, var dict) = rewriteType(code);
            var pre = GetCode("", "Structure", "Public ", "Public ");
            Helper.AssertCode(pre, act);

            //var v1 = (" ".Length + "Type".Length);
            var diff = "Structure".Length - "Type".Length;
            var menS = 0;
            var menL = "Public ".Length;
            var predict = new Dictionary<int, (int, int)> {
                {1, ("   ".Length + "Type".Length, diff)},
                {4, ("    End ".Length + "Type".Length, diff)},
                {2, (menS, menL)},
                {3, (menS, menL)}
            };
            predict.All(x => dict.Contains(x));
            foreach (var item in predict) {
                Assert.Equal(item.Value, dict[item.Key]);
            }
        }

        [Fact]
        public void TestPublicType() {
            var code = GetCode("Public", "Type", "", "");
            (var act, var dict) = rewriteType(code);
            var pre = GetCode("Public", "Structure", "Public ", "Public ");
            Helper.AssertCode(pre, act);

            var diff = "Structure".Length - "Type".Length;
            var menS = 0;
            var menL = "Public ".Length;
            var predict = new Dictionary<int, (int, int)> {
                {1, ("  Public ".Length + "Type".Length, diff)},
                {4, ("    End ".Length + "Type".Length, diff)},
                {2, (menS, menL)},
                {3, (menS, menL)}
            };
            predict.All(x => dict.Contains(x));
            foreach (var item in predict) {
                Assert.Equal(item.Value, dict[item.Key]);
            }
        }

        [Fact]
        public void TestPrivateType() {
            var code = GetCode("Private", "Type", "", "");
            (var act, var dict) = rewriteType(code);
            var pre = GetCode("Private", "Structure", "Public ", "Public ");
            Helper.AssertCode(pre, act);
        }

        [Fact]
        public void TestPublicTypePublicX() {
            var code = GetCode("Public", "Type", "Public", "");
            (var act, var dict) = rewriteType(code);
            var pre = GetCode("Public", "Structure", "Public Public", "Public ");
            Helper.AssertCode(pre, act);
        }

        [Fact]
        public void TestTypePublicY() {
            var code = GetCode("Public", "Type", "", "Public");
            (var act, var dict) = rewriteType(code);
            var pre = GetCode("Public", "Structure", "Public ", "Public Public");
            Helper.AssertCode(pre, act);
        }

        [Fact]
        public void TestPublicTypePublicXPublicY() {
            var code = GetCode("Public", "Type", "Public", "Public");
            //var act = rewriteType(code);
            (var act, var dict) = rewriteType(code);
            var pre = GetCode("Public", "Structure", "Public Public", "Public Public");
            Helper.AssertCode(pre, act);

            var diff = "Structure".Length - "Type".Length;
            var menS = 0;
            var menL = "Public ".Length;
            var predict = new Dictionary<int, (int, int)> {
                {1, ("  Public ".Length + "Type".Length, diff)},
                {4, ("    End ".Length + "Type".Length, diff)},
                {2, (menS, menL)},
                {3, (menS, menL)}
            };
            predict.All(x => dict.Contains(x));
            foreach (var item in predict) {
                Assert.Equal(item.Value, dict[item.Key]);
            }
        }

        private string getError(string valSt, string valX, string ValY, string end) {
            return @$"Public Module Module1
{valSt} Type Point
{valX} X As Long
{ValY} Y As Long
{end} End Type
End Module";
        }

        [Fact]
        public void TestErrorType() {
			var code = getError("a", "", "", "");
			(var act, var dict) = rewriteType(code);
            Helper.AssertCode(@$"Public Module Module1
a Type Point
 X As Long
 Y As Long
 End Type
End Module", act);
        }

        [Fact]
        public void TestErrorX() {
            var code = getError("", "a", "", "");
            (var act, var dict) = rewriteType(code);
            Helper.AssertCode(@$"Public Module Module1
 Structure Point
a X As Long
Public  Y As Long
 End Structure
End Module", act);
        }

        [Fact]
        public void TestErrorY() {
            var code = getError("", "", "a", "");
            (var act, var dict) = rewriteType(code);
            Helper.AssertCode(@$"Public Module Module1
 Structure Point
Public  X As Long
a Y As Long
 End Structure
End Module", act);
        }

        [Fact]
        public void TestErrorEnd() {
            var code = getError("", "", "", "a");
            (var act, var dict) = rewriteType(code);
            Helper.AssertCode(@$"Public Module Module1
 Structure Point
Public  X As Long
Public  Y As Long
Public a End Structure
End Module", act);
        }
    }
}
