using VBACodeAnalysis;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace TestProject {
	public class TestRewriteAnnotationAs {

        private string GetCode() {
            return @$"Public Module Module1
' @As Test
Dim obj As Variant
' @As Test2
Dim obj2 As Variant
' @As   
Dim obj3 As Variant
Dim obj4 As Variant
End Module";
        }
        private string GetPreCode() {
            return @$"Public Module Module1
' @As Test
Dim obj As Test
' @As Test2
Dim obj2 As Test2
' @As   
Dim obj3 As Variant
Dim obj4 As Variant
End Module";
        }

        string RewriteAs(string code) {
            var doc = Helper.MakeDoc(code);
            var rew = new Rewrite(new RewriteSetting());
            var root = rew.AnnotationAs(doc.GetSyntaxRootAsync().Result);
            return root.ToString();
        }

        [Fact]
        public void TestRewriteAs() {
            var code = GetCode();
            var act = RewriteAs(code);
            var pre = GetPreCode();
            Helper.AssertCode(pre, act);
        }
    }
}
