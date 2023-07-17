using VBACodeAnalysis;
using System.Collections.Generic;
using Xunit;

namespace TestProject {
    public class TestRewritePredefined {
        Rewrite rewrite;

        private void Setup() {
            var setting = new RewriteSetting();
            setting.VBAPredefined.ModuleName = "VBAFunction";
            rewrite = new Rewrite(setting);
        }

        private string GetCode(string prefix) {
            return $@"
Public Class Cells
    Public Sub test1()
        Dim v As Object
        v =    {prefix}CStr(""a"" & {prefix}CStr({prefix}CStr(1)))   'aaa
        v = Trim(""a"")
    End Sub
End Class";
        }

        [Fact]
        public void TestRewrite() {
            Setup();

            var doc = Helper.MakeDoc(GetCode(""));
            var docRoot = doc.GetSyntaxRootAsync().Result;
            var result = rewrite.Predefined(docRoot);
            var actCode = result.root.GetText().ToString();
            var preCode = GetCode("VBAFunction.");
            Helper.AssertCode(preCode, actCode);

            var predict = new Dictionary<int, List<LocationDiff>> {
                {4, new List<LocationDiff>{ 
                    new LocationDiff(4, 15, 12),
                    new LocationDiff(4, 26, 12),
                    new LocationDiff(4, 31, 12),
                }}
            };
            Helper.AssertLocationDiffDict(predict, result.dict);
        }
    }
}
