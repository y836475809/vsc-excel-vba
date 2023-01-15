using System;
using System.Text;

namespace ConsoleAppEncode {
    class Program {
        static void Main(string[] args) {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var sjis_enc = Encoding.GetEncoding("SHIFT_JIS");
            var utf8_enc = new UTF8Encoding(false);

            var sjis_text1 = Helper.readFile("test_module1.bas", sjis_enc);
            var utf8_text1 = Helper.ConvertEncoding2(sjis_text1, sjis_enc, utf8_enc);
            Helper.writeFile("test_module1-utf8.bas", utf8_text1, utf8_enc);

            var utf8_text2 = Helper.readFile("test_module1-utf8.bas", utf8_enc);
            var sjis_text2 = Helper.ConvertEncoding2(utf8_text2, utf8_enc, sjis_enc);
            Helper.writeFile("test_module1-sjis.bas", sjis_text2, sjis_enc);

            var m = 0;
        }
    }
}
