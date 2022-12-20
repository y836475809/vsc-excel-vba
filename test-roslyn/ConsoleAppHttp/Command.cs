using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleAppHttp {
    class Command {
        public string id { get; set; }
        public string json_string { get; set; }
    }

    class ResponseCompletion {
        public List<string> items { get; set; }
    }
}
