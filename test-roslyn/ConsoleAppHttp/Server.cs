using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Text.Json;

namespace ConsoleAppServer
{
    public class Server
    {
        public event EventHandler<DocumentAddedEventArgs> DocumentAdded;
        public event EventHandler<DocumentChangedEventArgs> DocumentChanged;
        public event EventHandler<CompletionEventArgs> CompletionReq;
        private HttpListener listener;

        public void Setup(int port)
        {
            listener = new HttpListener();

            // リスナー設定
            listener.Prefixes.Clear();
            listener.Prefixes.Add(@$"http://localhost:{port}/");
            listener.Start();
        }

        public void Run()
        {
            bool run = true;
            while (run)
            {
                // リクエスト取得
                var context = listener.GetContext();
                var request = context.Request;
                Command cmd;
                using (var reader = new StreamReader(request.InputStream, request.ContentEncoding))
                {
                    var json_str = reader.ReadToEnd();
                    cmd = JsonSerializer.Deserialize<Command>(json_str);
                    //Console.WriteLine(json_str);
                }

                // レスポンス取得
                var response = context.Response;
                switch (cmd?.Id)
                {
                    case "AddDocument":
                        DocumentAdded?.Invoke(this, new DocumentAddedEventArgs(cmd.FilePath));
                        Response(response, 202);
                        break;
                    case "ChangeDocument":
                        DocumentChanged?.Invoke(this, new DocumentChangedEventArgs(cmd.FilePath, cmd.Text));
                        Response(response, 202);
                        break;
                    case "Completion":
                        var args = new CompletionEventArgs(cmd.FilePath, cmd.Text, cmd.Position);
                        CompletionReq?.Invoke(this, args);
                        Response(response, args.Items);
                        break;
                    //case "Exit":
                    case "Shutdown":
                        run = false;
                        break;
                }
                response.Close();


                //// HTMLを表示する
                //if (request != null)
                //{
                //    var res_com = new ResponseCompletion();
                //    res_com.items = new List<string>() { "aaa", "bbb" };
                //    var text = Encoding.UTF8.GetBytes(JsonSerializer.Serialize<ResponseCompletion>(res_com));
                //    response.ContentType = "application/json";
                //    response.ContentLength64 = text.Length;
                //    response.OutputStream.Write(text, 0, text.Length);
                //}
                //else
                //{
                //    response.StatusCode = 404;
                //}
                //response.Close();
            }

        }

        private void Response(HttpListenerResponse response, int StatusCode) {
            response.StatusCode = StatusCode;
            //response.Close();
        }
        private void Response(HttpListenerResponse response, List<string> CompletionItems)
        {
            var res_com = new ResponseCompletion();
            res_com.items = CompletionItems;
            var text = Encoding.UTF8.GetBytes(JsonSerializer.Serialize<ResponseCompletion>(res_com));
            response.ContentType = "application/json";
            response.ContentLength64 = text.Length;
            response.OutputStream.Write(text, 0, text.Length);
            //response.Close();
        }
    }
}
