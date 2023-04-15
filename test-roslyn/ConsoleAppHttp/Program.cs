namespace ConsoleAppServer {
	class Program {
        static void Main(string[] args) {
            var app = new App();
            app.Initialize();
            app.Run(9088);
        }
    }
}
