using ConsoleApp1;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json;

namespace ConsoleAppServer {
	class App {
		private CodeAdapter codeAdapter;
		private Server server;
		private MyCodeAnalysis mc;

		public App() {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        public void Run(int port) {
            server.Setup(port);
            server.Run();
        }

        private void Reset() {
            codeAdapter = new CodeAdapter();
            mc = new MyCodeAnalysis();
            var settings = LoadConfig();
            mc.setSetting(settings.RewriteSetting);
        }

        public void Initialize() {  
            Reset();

            server = new Server();
            server.ResetReq += (object sender, EventArgs e) => {
                mc = new MyCodeAnalysis();
                codeAdapter = new CodeAdapter();
            };
            server.DocumentAdded += (object sender, DocumentAddedEventArgs e) => {
                foreach (var FilePath in e.FilePaths) {
                    codeAdapter.SetCode(FilePath, Helper.getCode(FilePath));
                    var vbCode = codeAdapter.GetVbCodeInfo(FilePath).VbCode;
                    mc.AddDocument(FilePath, vbCode);
                }
            };
            server.DocumentDeleted += (object sender, DocumentDeletedEventArgs e) => {
                foreach (var FilePath in e.FilePaths) {
                    mc.DeleteDocument(FilePath);
                    codeAdapter.Delete(FilePath);
                }
            };
            server.DocumentRenamed += (object sender, DocumentRenamedEventArgs e) => {
                mc.DeleteDocument(e.OldFilePath);
                codeAdapter.Delete(e.OldFilePath);
                codeAdapter.SetCode(e.NewFilePath, Helper.getCode(e.NewFilePath));
                var vbCode = codeAdapter.GetVbCodeInfo(e.NewFilePath).VbCode;
                mc.AddDocument(e.NewFilePath, vbCode);

            };
            server.DocumentChanged += (object sender, DocumentChangedEventArgs e) => {
                codeAdapter.SetCode(e.FilePath, e.Text);
                var vbCode = codeAdapter.GetVbCodeInfo(e.FilePath).VbCode;
                mc.ChangeDocument(e.FilePath, vbCode);
            };
            server.CompletionReq += async (object sender, CompletionEventArgs e) => {
                var vbCodeInfo = codeAdapter.GetVbCodeInfo(e.FilePath);
                var vbCode = vbCodeInfo.VbCode;
                var posOffset = vbCodeInfo.PositionOffset;
                var line = e.Line - vbCodeInfo.LineOffset;
                if (line < 0) {
                    e.Items = new List<CompletionItem>();
                    return;
                }
                var Items = await mc.GetCompletions(e.FilePath, vbCode, line, e.Chara);
                e.Items = Items;
            };
            server.DefinitionReq += async (object sender, DefinitionEventArgs e) => {
                var list = new List<DefinitionItem>();
                var vbCodeInfo = codeAdapter.GetVbCodeInfo(e.FilePath);
                var vbCode = vbCodeInfo.VbCode;
                var posOffset = vbCodeInfo.PositionOffset;
                var line = e.Line - vbCodeInfo.LineOffset;
                if (line < 0) {
                    e.Items = list;
                    return;
                }
                var Items = await mc.GetDefinitions(e.FilePath, vbCode, line, e.Chara);
                Location adjustLocation(Location location, int lineOffset, int chaOffset) {
                    location.Line += lineOffset;
                    location.Positon += chaOffset;
                    if (location.Line < 0) {
                        location.Line = 0;
                    }
                    if (location.Positon < 0) {
                        location.Positon = 0;
                    }
                    return new Location(location.Positon, location.Line, location.Character);
                }
                foreach (var item in Items) {
                    var itemVbCodeInfo = codeAdapter.GetVbCodeInfo(item.FilePath);
                    var itemLineOffset = itemVbCodeInfo.LineOffset;
                    var itemPosOffset = itemVbCodeInfo.PositionOffset;
                    if (item.IsKindClass()) {
                        item.Start.Positon = 0;
                        item.Start.Line = 0;
                        item.Start.Character = 0;
                        item.End.Positon = 0;
                        item.End.Line = 0;
                        item.End.Character = 0;
                    } else {
                        item.Start = adjustLocation(item.Start, itemLineOffset, itemPosOffset);
                        item.End = adjustLocation(item.End, itemLineOffset, itemPosOffset);
                    }
                    list.Add(item);
                }
                e.Items = list;
            };
            server.HoverReq += async (object sender, CompletionEventArgs e) => {
                var list = new List<CompletionItem>();
                var vbCodeInfo = codeAdapter.GetVbCodeInfo(e.FilePath);
                var vbCode = vbCodeInfo.VbCode;
                var posOffset = vbCodeInfo.PositionOffset;
                var line = e.Line - vbCodeInfo.LineOffset;
                if (line < 0) {
                    e.Items = list;
                    return;
                }
                var Items = await mc.GetDefinitions(e.FilePath, vbCode, line, e.Chara);
                foreach (var item in Items) {
                    var sp = item.Start.Positon;
                    var ep = item.End.Positon;
                    var hoverItem = await mc.GetHover(item.FilePath, e.Text, (int)((sp + ep) / 2));
                    list.Add(hoverItem);
                }
                e.Items = list;
            };
            server.DiagnosticReq += async (object sender, DiagnosticEventArgs e) => {
                var vbCodeInfo = codeAdapter.GetVbCodeInfo(e.FilePath);
                var lineOffset = vbCodeInfo.LineOffset;
                var items = await mc.GetDiagnostics(e.FilePath);
                foreach (var item in items) {
                    item.StartLine += lineOffset;
                    item.EndLine += lineOffset;
                }
                e.Items = items;
            };
            server.DebugGetDocumentsEvent += (object sender, DebugEventArgs e) => {
                e.Text = JsonSerializer.Serialize(codeAdapter.getVbCodeDict());
            };
        }
        private Settings LoadConfig() {
            var builder = new ConfigurationBuilder();
            builder.AddJsonFile("app.json");
            var config = builder.Build();
            var settings = new Settings();
            config.GetSection("App").Bind(settings);
            settings.convert();
            return settings;
        }
    }
}
