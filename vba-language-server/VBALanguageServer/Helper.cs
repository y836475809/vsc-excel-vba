using System.IO;
using System.Text;
using System.Text.Json;

namespace VBALanguageServer {
    class Helper {
        public static string getCode(string filePath) {
            var enc = Encoding.GetEncoding("shift_jis");
            using (var sr = new StreamReader(filePath, enc)) {
                return sr.ReadToEnd();
            }
        }
    }
}
