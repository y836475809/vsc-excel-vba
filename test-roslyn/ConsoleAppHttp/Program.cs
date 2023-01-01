using ConsoleApp1;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Text.Json;

namespace ConsoleAppServer {
    class Program {
        static void Main(string[] args) {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var docCache = new Dictionary<string, string>();
            
            var mc = new MyCodeAnalysis();
            var server = new Server();
            server.DocumentAdded += (object sender, DocumentAddedEventArgs e) => {
                //var code = "";
                if (!docCache.ContainsKey(e.FilePath)) {
                    docCache[e.FilePath] = Helper.getCode(e.FilePath);
                }
                var code = docCache[e.FilePath];
                e.Text = code;
                mc.AddDocument(e.FilePath, code);
            };
            server.DocumentChanged += (object sender, DocumentChangedEventArgs e) => {
                //changedDocSet.Add(e.FilePath);
                //var code = Helper.getCode(e.FilePath);
                //mc.AddDocument(e.FilePath, code);
                docCache[e.FilePath] = e.Text;
                mc.ChangeDocument(e.FilePath, e.Text);
            };
            server.CompletionReq += async (object sender, CompletionEventArgs e) => {
                var Items = await mc.GetCompletions(e.FilePath, e.Text, e.Position);
                e.Items = Items;
            };
            server.DefinitionReq += async (object sender, DefinitionEventArgs e) => {
                var Items = await mc.GetDefinitions(e.FilePath, e.Text, e.Position);
                var list = new List<DefinitionItem>();
                foreach (var item in Items) {
                    list.Add(item);
                }
                e.Items = list;
            };
            server.HoverReq += async (object sender, CompletionEventArgs e) => {
                var list = new List<CompletionItem>();
                var Items = await mc.GetDefinitions(e.FilePath, e.Text, e.Position);
                foreach (var item in Items) {
                    if (!docCache.ContainsKey(item.FilePath)) {
                        docCache[e.FilePath] = Helper.getCode(item.FilePath);
                    }
                    var code = docCache[e.FilePath];
                    list.Add(await mc.GetHover(item.FilePath, code, (int)((item.Start + item.End)/2)));
                }
                e.Items = list;
            };
            server.Setup(9088);
            server.Run();

            //try {
            //    // HTTPリスナー作成
            //    var listener = new HttpListener();

            //    // リスナー設定
            //    listener.Prefixes.Clear();
            //    listener.Prefixes.Add(@"http://localhost:9088/");

            //    // リスナー開始
            //    listener.Start();

            //    while (true) {
            //        // リクエスト取得
            //        var context = listener.GetContext();
            //        var request = context.Request;

            //        using (var reader = new StreamReader(request.InputStream, request.ContentEncoding)) {
            //            var json_str = reader.ReadToEnd();
            //            var cmd = JsonSerializer.Deserialize<Command>(json_str);
            //            Console.WriteLine(json_str);
            //        }

            //        // レスポンス取得
            //        var response = context.Response;

            //        // HTMLを表示する
            //        if (request != null) {
            //            var res_com = new ResponseCompletion();
            //            res_com.items = new List<string>() { "aaa", "bbb" };
            //            var text = Encoding.UTF8.GetBytes(JsonSerializer.Serialize<ResponseCompletion>(res_com));
            //            response.ContentType = "application/json";
            //            response.ContentLength64 = text.Length;
            //            response.OutputStream.Write(text, 0, text.Length);
            //        } else {
            //            response.StatusCode = 404;
            //        }
            //        response.Close();
            //    }

            //} catch (Exception e) {
            //    Console.WriteLine(e.Message);
            //}
        }

        private static void Server_HoverReq(object sender, CompletionEventArgs e) {
            throw new NotImplementedException();
        }

        private static void Server_DefinitionReq(object sender, DefinitionEventArgs e) {
            throw new NotImplementedException();
        }
    }
}
