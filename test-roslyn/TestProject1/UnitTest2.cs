using ConsoleApp1;
using System;
using System.Collections.Generic;
using System.Text;
using Xunit;

namespace TestProject1 {
    public class UnitTest2 {
        [Fact]
        public void TestParseDescriptionXML() {
            var mc = new MyCodeAnalysis();
            string str =
@"
<member name=""M: Person.SayHello(System.Object, System.Object)"">
    <summary>
        テストメッセージ
    </summary>
    <param name = ""val1"">引数1</param>
    <param name = ""val2"">引数1</param>
    <returns>True</returns>
</member>
";
            var d = mc.ParseDescriptionXML(str);
            var m = 0;
        }
    }
}
