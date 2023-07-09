using VBALanguageServer;
using Xunit;

namespace TestProject {
    public class TestVBACode {

        [Fact]
        public void TestModuleCodeAdapter() {
            string code =
@"
Attribute VB_Name = ""m1""

''' <summary>
'''  モジュールbuf
''' </summary>
Private buf As String
";
            var cd = new CodeAdapter();
            cd.parse("m1.bas", code, out VbCodeInfo vbCodeInfo);
            Assert.Equal(1, vbCodeInfo.LineOffset);
        }

        [Fact]
        public void TestClassCodeAdapter() {
            string code =
@"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""Person""
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit On
''' <summary>
'''  メンバ変数
''' </summary> 
Public Name As String
";
            var cd = new CodeAdapter();
            cd.parse("m1.cls", code, out VbCodeInfo vbCodeInfo);
            Assert.Equal(9, vbCodeInfo.LineOffset);
        }

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
            Assert.Equal(9, vbCodeInfo.LineOffset);
            var pre = MakeClassCode("", "", "Public Class Test", "\r\nEnd Class");
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
			Assert.Equal(8, vbCodeInfo.LineOffset);
            var pre = MakeClassCode("Public Class Test", "\r\n", value, "\r\nEnd Class");
            Helper.AssertCode(pre, vbCodeInfo.VbCode);
		}
	}
}