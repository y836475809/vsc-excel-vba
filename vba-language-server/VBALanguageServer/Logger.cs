using System;
using System.Collections.Generic;
using System.Text;

namespace VBALanguageServer {
	class Logger {
        public void Info(string msg) {
            Write("Info", msg);
        }

        private void Write(string Level, string msg) {
            var date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            Console.WriteLine($"[{date}][{Level}] {msg}");
		}
    }
}
