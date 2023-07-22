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

        public class LowerCaseNamingPolicy : JsonNamingPolicy {
            public override string ConvertName(string name) =>
                name.ToLower();
        }

        public static JsonSerializerOptions getJsonOptions() {
            var options = new JsonSerializerOptions {
                PropertyNamingPolicy = new LowerCaseNamingPolicy(),
                WriteIndented = true
            };
            return options;
        }
    }
}
