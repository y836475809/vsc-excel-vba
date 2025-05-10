using System;
using System.Collections.Generic;
using System.Text;

namespace VBALanguageServer {
	static class Logger {
        public static void Info(string msg) {
            Write("Info", msg);
        }

		public static void Error(string msg) {
			Write("Error", msg);
		}

		private static void Write(string Level, string msg) {
            var date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            Console.Error.WriteLine($"[{date}][{Level}] {msg}");
		}
    }
}
