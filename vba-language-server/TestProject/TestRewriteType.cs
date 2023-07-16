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

        (string, Dictionary<int, List<LocationDiff>>) rewriteType(string code) {
            var doc = Helper.MakeDoc(code);
            var rew = new Rewrite(new RewriteSetting());
            var result = rew.TypeStatement(doc.GetSyntaxRootAsync().Result);
            return (result.root.GetText().ToString(), result.dict);
        }

        [Fact]
        public void TestType() {
            var code = GetCode("", "Type", "", "");
            (var act, var dict) = rewriteType(code);
            var pre = GetCode("", "Structure", " Public", " Public");
            Helper.AssertCode(pre, act);

            var diff = "Structure".Length - "Type".Length;
            var menS = 1;
            var menL = "Public ".Length;
            var predict = new Dictionary<int, List<LocationDiff>> {
                {1, new List<LocationDiff>{ new LocationDiff(1, "   ".Length + "Type".Length, diff)} },
                {4, new List<LocationDiff>{ new LocationDiff(4, "    End ".Length + "Type".Length, diff)} },
                {2, new List<LocationDiff>{ new LocationDiff(2, menS, menL)} },
                {3, new List<LocationDiff>{ new LocationDiff(3, menS, menL)} },
            };
            Helper.AssertLocationDiffDict(predict, dict);
        }

        [Fact]
        public void TestPublicType() {
            var code = GetCode("Public", "Type", "", "");
            (var act, var dict) = rewriteType(code);
            var pre = GetCode("Public", "Structure", " Public", " Public");
            Helper.AssertCode(pre, act);

            var diff = "Structure".Length - "Type".Length;
            var menS = 1;
            var menL = "Public ".Length;
            var predict = new Dictionary<int, List<LocationDiff>> {
                {1, new List<LocationDiff>{ new LocationDiff(1, "  Public ".Length + "Type".Length, diff)} },
                {4, new List<LocationDiff>{ new LocationDiff(4, "    End ".Length + "Type".Length, diff)} },
                {2, new List<LocationDiff>{ new LocationDiff(2, menS, menL)} },
                {3, new List<LocationDiff>{ new LocationDiff(3, menS, menL)} },
            };
            Helper.AssertLocationDiffDict(predict, dict);
        }

        [Fact]
        public void TestPrivateType() {
            var code = GetCode("Private", "Type", "", "");
            (var act, var dict) = rewriteType(code);
            var pre = GetCode("Private", "Structure", " Public", " Public");
            Helper.AssertCode(pre, act);
        }

        [Fact]
        public void TestPublicTypePublicX() {
            var code = GetCode("Public", "Type", "Public", "");
            (var act, var dict) = rewriteType(code);
            var pre = GetCode("Public", "Structure", "Public Public", " Public");
            Helper.AssertCode(pre, act);
        }

        [Fact]
        public void TestTypePublicY() {
            var code = GetCode("Public", "Type", "", "Public");
            (var act, var dict) = rewriteType(code);
            var pre = GetCode("Public", "Structure", " Public", "Public Public");
            Helper.AssertCode(pre, act);
        }

        [Fact]
        public void TestPublicTypePublicXPublicY() {
            var code = GetCode("Public", "Type", "Public", "Public");
            (var act, var dict) = rewriteType(code);
            var pre = GetCode("Public", "Structure", "Public Public", "Public Public");
            Helper.AssertCode(pre, act);

            var diff = "Structure".Length - "Type".Length;
            var menS = 0;
            var menL = "Public ".Length;
            var predict = new Dictionary<int, List<LocationDiff>> {
                {1, new List<LocationDiff>{ new LocationDiff(1, "  Public ".Length + "Type".Length, diff)} },
                {4, new List<LocationDiff>{ new LocationDiff(4, "    End ".Length + "Type".Length, diff)} },
                {2, new List<LocationDiff>{ new LocationDiff(2, menS, menL)} },
                {3, new List<LocationDiff>{ new LocationDiff(3, menS, menL)} },
            };
            Helper.AssertLocationDiffDict(predict, dict);
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
 Public Y As Long
 End Structure
End Module", act);
        }

        [Fact]
        public void TestErrorY() {
            var code = getError("", "", "a", "");
            (var act, var dict) = rewriteType(code);
            Helper.AssertCode(@$"Public Module Module1
 Structure Point
 Public X As Long
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
 Public X As Long
 Public Y As Long
Public a End Structure
End Module", act);
        }

        [Fact]
        public void TestPublicType2() {
            var code = @$"Public Module Module1
Type Point
' test1
  X As Long

  Y    As Long
  
End Type
End Module";
            (var act, var dict) = rewriteType(code);
            var pre = @$"Public Module Module1
Structure Point
' test1
  Public X As Long

  Public Y    As Long
  
End Structure
End Module";
            Helper.AssertCode(pre, act);

            var diff = "Structure".Length - "Type".Length;
            var menS = 2;
            var menL = "Public ".Length;
            var predict = new Dictionary<int, List<LocationDiff>> {
                {1, new List<LocationDiff>{ new LocationDiff(1, "Type".Length, diff)} },
                {7, new List<LocationDiff>{ new LocationDiff(7, "End ".Length + "Type".Length, diff)} },
                {3, new List<LocationDiff>{ new LocationDiff(3, menS, menL)} },
                {5, new List<LocationDiff>{ new LocationDiff(5, menS, menL)} },
            };
            Helper.AssertLocationDiffDict(predict, dict);
        }
    }
}
