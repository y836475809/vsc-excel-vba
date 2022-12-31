using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;

namespace ConsoleAppServer {
    class Helper {
        //public static string getPath(string fileName, [CallerFilePath] string filePath = "") {
        //    return Path.Combine(Path.GetDirectoryName(filePath), "code", fileName);
        //}
        //public static int getPosition(string code, string target) {
        //    return code.LastIndexOf(target) + target.Length;
        //}
        public static string getCode(string filePath) {
            using (var sr = new StreamReader(filePath, Encoding.GetEncoding("shift_jis"))) {
                return sr.ReadToEnd();
            }
        }
    }
}
