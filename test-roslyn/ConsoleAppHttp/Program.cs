﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Text.Json;

namespace ConsoleAppHttp {
    class Program {
        static void Main(string[] args) {
            try {
                // HTTPリスナー作成
                var listener = new HttpListener();

                // リスナー設定
                listener.Prefixes.Clear();
                listener.Prefixes.Add(@"http://localhost:9088/");

                // リスナー開始
                listener.Start();

                while (true) {
                    // リクエスト取得
                    var context = listener.GetContext();
                    var request = context.Request;

                    using (var reader = new StreamReader(request.InputStream, request.ContentEncoding)) {
                        var json_str = reader.ReadToEnd();
                        var cmd = JsonSerializer.Deserialize<Command>(json_str);
                        Console.WriteLine(json_str);
                    }

                    // レスポンス取得
                    var response = context.Response;

                    // HTMLを表示する
                    if (request != null) {
                        var res_com = new ResponseCompletion();
                        res_com.items = new List<string>() { "aaa", "bbb" };
                        var text = Encoding.UTF8.GetBytes(JsonSerializer.Serialize<ResponseCompletion>(res_com));
                        response.ContentType = "application/json";
                        response.ContentLength64 = text.Length;
                        response.OutputStream.Write(text, 0, text.Length);
                    } else {
                        response.StatusCode = 404;
                    }
                    response.Close();
                }

            } catch (Exception e) {
                Console.WriteLine(e.Message);
            }
        }
    }
}