using ConsoleAppServer;
using Xunit;

namespace TestProject1 {
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
    }
}