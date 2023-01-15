using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;

namespace ConsoleAppEncode {
    class Helper {
        public static string getPath(string fileName, [CallerFilePath] string filePath = "") {
            return Path.Combine(Path.GetDirectoryName(filePath), "code", fileName);
        }

        public static string readFile(string fileName, Encoding enc) {
            var filePath = Helper.getPath(fileName);
            using (var sr = new StreamReader(filePath, enc)) {
                return sr.ReadToEnd();
            }
        }

        public static void writeFile(string fileName, string text, Encoding enc) {
            var filePath = Helper.getPath(fileName);
            using (var sw = new StreamWriter(filePath, false, enc)) {
                sw.WriteAsync(text);
            }
        }

        public static string ConvertEncoding(string src, Encoding destEnc) {
            byte[] src_temp = Encoding.GetEncoding("SHIFT_JIS").GetBytes(src);
            byte[] dest_temp = Encoding.Convert(Encoding.GetEncoding("SHIFT_JIS"), destEnc, src_temp);
            string ret = destEnc.GetString(dest_temp);
            return ret;
        }

        public static string ConvertEncoding2(string src, Encoding srcEnc, Encoding destEnc) {
            byte[] src_temp = srcEnc.GetBytes(src);
            byte[] dest_temp = Encoding.Convert(srcEnc, destEnc, src_temp);
            string ret = destEnc.GetString(dest_temp);
            return ret;
        }
    }
}
