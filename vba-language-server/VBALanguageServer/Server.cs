using System;
using System.IO;
using System.Net;
using System.Text;
using System.Text.Json;

namespace VBALanguageServer {
    public class Server
    {
        public event EventHandler<DocumentAddedEventArgs> DocumentAdded;
        public event EventHandler<DocumentDeletedEventArgs> DocumentDeleted;
        public event EventHandler<DocumentRenamedEventArgs> DocumentRenamed;
        public event EventHandler<DocumentChangedEventArgs> DocumentChanged;
        public event EventHandler<CompletionEventArgs> CompletionReq;
        public event EventHandler<DefinitionEventArgs> DefinitionReq;
        public event EventHandler<CompletionEventArgs> HoverReq;
        public event EventHandler<DiagnosticEventArgs> DiagnosticReq;
        public event EventHandler<EventArgs> ResetReq;
        public event EventHandler<DebugEventArgs> DebugGetDocumentsEvent;
        public event EventHandler<ReferencesEventArgs> ReferencesReq;
        public event EventHandler<SignatureHelpEventArgs> SignatureHelpReq;
        private HttpListener listener;
        private JsonSerializerOptions jsonOptions;

        private const int ResponseOK = 200;

        public Server() {
            jsonOptions = Helper.getJsonOptions();
        }

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
            bool ignoreShutdown = false;
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
                }

                // レスポンス取得
                var response = context.Response;
				try {
                    switch (cmd?.id) {
                        case "IsReady":
                            Response(response, ResponseOK);
                            break;
                        case "AddDocuments":
                            var args_add = new DocumentAddedEventArgs(cmd.filepaths);
                            DocumentAdded?.Invoke(this, args_add);
                            Response(response, ResponseOK);
                            break;
                        case "DeleteDocuments":
                            var args_del = new DocumentDeletedEventArgs(cmd.filepaths);
                            DocumentDeleted?.Invoke(this, args_del);
                            Response(response, ResponseOK);
                            break;
                        case "RenameDocument":
                            var args_rename = new DocumentRenamedEventArgs(cmd.filepaths[0], cmd.filepaths[1]);
                            DocumentRenamed?.Invoke(this, args_rename);
                            Response(response, ResponseOK);
                            break;
                        case "ChangeDocument":
                            DocumentChanged?.Invoke(this, new DocumentChangedEventArgs(cmd.filepaths[0], cmd.text));
                            Response(response, ResponseOK);
                            break;
                        case "Completion":
                            var args = new CompletionEventArgs(cmd.filepaths[0], cmd.text, cmd.line, cmd.chara);
                            CompletionReq?.Invoke(this, args);
                            Response(response, args.Items);
                            break;
                        case "Definition":
                            var args_def = new DefinitionEventArgs(cmd.filepaths[0], cmd.text, cmd.line, cmd.chara);
                            DefinitionReq?.Invoke(this, args_def);
                            Response(response, args_def.Items);
                            break;
                        case "Hover":
                            var args_hover = new CompletionEventArgs(cmd.filepaths[0], cmd.text, cmd.line, cmd.chara);
                            HoverReq?.Invoke(this, args_hover);
                            Response(response, args_hover.Items);
                            break;
                        case "Diagnostic":
                            var args_diagnostic = new DiagnosticEventArgs(cmd.filepaths[0]);
                            DiagnosticReq?.Invoke(this, args_diagnostic);
                            Response(response, args_diagnostic.Items);
                            break;
                        case "Shutdown":
                            if (!ignoreShutdown) {
                                run = false;
                            }
                            break;
                        case "Reset":
                            ResetReq?.Invoke(this, new EventArgs());
                            Response(response, ResponseOK);
                            break;
                        case "IgnoreShutdown":
                            ignoreShutdown = true;
                            Response(response, ResponseOK);
                            break;
                        case "Debug:GetDocuments":
                            var args_debug = new DebugEventArgs();
                            string debugInfo = string.Empty;
                            DebugGetDocumentsEvent?.Invoke(this, args_debug);
                            Response(response, args_debug.Text);
                            break;
                        case "References":
                            var args_refs = new ReferencesEventArgs(cmd.filepaths[0], cmd.line, cmd.chara);
                            ReferencesReq?.Invoke(this, args_refs);
                            Response(response, args_refs.Items);
                            break;
                        case "SignatureHelp":
                            var argsSignatureHelp = new SignatureHelpEventArgs(cmd.filepaths[0], cmd.text, cmd.line, cmd.chara);
                            SignatureHelpReq?.Invoke(this, argsSignatureHelp);
                            Response(response, argsSignatureHelp.Items);
                            break;
                    }
                } catch (Exception e) {
                    response.StatusDescription = e.Message + ": " + e.StackTrace;
                    Response(response, 500); ;
				}
                if (run) {
                    response.Close();
                }
            }
        }

        private void Response(HttpListenerResponse response, int StatusCode) {
            response.StatusCode = StatusCode;
        }

        private void Response<T>(HttpListenerResponse response, T CompletionItems) {
            var text = Encoding.UTF8.GetBytes(
                JsonSerializer.Serialize<T>(CompletionItems, jsonOptions));
            response.ContentType = "application/json";
            response.ContentLength64 = text.Length;
            response.OutputStream.Write(text, 0, text.Length);
            response.StatusCode = ResponseOK;
        }
    }
}
