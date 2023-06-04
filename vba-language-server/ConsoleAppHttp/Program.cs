using System;

namespace ConsoleAppServer {
	class Program {
        static void Main(string[] args) {
            if (args.Length == 0) {
                Console.WriteLine($"Requires port as args");
                return;
            }
            var port = int.Parse(args[0]);
            Console.WriteLine($"port={port}");
            var app = new App();
            app.Initialize();
            app.Run(port);
        }
    }
}
