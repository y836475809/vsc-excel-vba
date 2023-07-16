using System.Linq;
using VBALanguageServer;
using Xunit;

namespace TestProject {
    public class TestVBACode {
        private string MakeVBAAttr() {
            return @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""Test""
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
";
        }
        private string MakeClassCode(string attr, 
            string preDec, string preVar, string postCode) {
            return
@$"{attr}
{preDec}
Dim PreName As String

{preVar}
Public Name As String
Publuc Sub Test()
End Sub{postCode}";
        }

        [Theory]
        [InlineData("' @class")]
        [InlineData("' @Class")]
        [InlineData("' @CLASS")]
        [InlineData("'@class")]
        [InlineData("    '   @class")]
        [InlineData("    '   @class  ")]
        [InlineData("' @class Test")]
        public void TestClassAnnotationValid(string value) {
            var vbacode = MakeClassCode(MakeVBAAttr(), "", value, "");
            var cd = new CodeAdapter();
            cd.parse("m1.cls", vbacode, out VbCodeInfo vbCodeInfo);
            var att = string.Concat(Enumerable.Repeat("\r\n", 9));
            var pre = MakeClassCode(att, "", "Public Class Test", "\r\nEnd Class");
            Helper.AssertCode(pre, vbCodeInfo.VbCode);
        }

        [Theory]
        [InlineData("' @classTest")]
        [InlineData("' @ class")]
        [InlineData("")]
        public void TestClassAnnotationInvalid(string value) {
            var code = MakeClassCode(MakeVBAAttr(), "", value, "");
            var cd = new CodeAdapter();
			cd.parse("m1.cls", code, out VbCodeInfo vbCodeInfo);
            var att = string.Concat(Enumerable.Repeat("\r\n", 8));
            var pre = MakeClassCode($"{att}Public Class Test", "\r\n", value, "\r\nEnd Class");
            Helper.AssertCode(pre, vbCodeInfo.VbCode);
		}
	}
}