using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;

namespace ConsoleAppServer {
    class Helper {
        //public static string getPath(string fileName, [CallerFilePath] string filePath = "") {
        //    return Path.Combine(Path.GetDirectoryName(filePath), "code", fileName);
        //}
        //public static int getPosition(string code, string target) {
        //    return code.LastIndexOf(target) + target.Length;
        //}
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
