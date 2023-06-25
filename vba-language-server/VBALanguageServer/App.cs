﻿﻿using VBACodeAnalysis;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;

namespace VBALanguageServer {
	class App {
		private CodeAdapter codeAdapter;
		private Server server;
		private MyCodeAnalysis mc;
        private Logger logger;

        public App() {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            logger = new Logger();
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
                Reset();
                logger.Info("ResetReq");
            };
            server.DocumentAdded += (object sender, DocumentAddedEventArgs e) => {
                foreach (var FilePath in e.FilePaths) {
                    codeAdapter.SetCode(FilePath, Helper.getCode(FilePath));
                    var vbCode = codeAdapter.GetVbCodeInfo(FilePath).VbCode;
                    mc.AddDocument(FilePath, vbCode, false);
                }
                mc.ApplyChanges(e.FilePaths);
                logger.Info("DocumentAdded");
            };
            server.DocumentDeleted += (object sender, DocumentDeletedEventArgs e) => {
                foreach (var FilePath in e.FilePaths) {
                    mc.DeleteDocument(FilePath);
                    codeAdapter.Delete(FilePath);
                }
                logger.Info("DocumentDeleted");
            };
            server.DocumentRenamed += (object sender, DocumentRenamedEventArgs e) => {
                mc.DeleteDocument(e.OldFilePath);
                codeAdapter.Delete(e.OldFilePath);
                codeAdapter.SetCode(e.NewFilePath, Helper.getCode(e.NewFilePath));
                var vbCode = codeAdapter.GetVbCodeInfo(e.NewFilePath).VbCode;
                mc.AddDocument(e.NewFilePath, vbCode);
                logger.Info("DocumentRenamed");
            };
            server.DocumentChanged += (object sender, DocumentChangedEventArgs e) => {
                codeAdapter.SetCode(e.FilePath, e.Text);
                var vbCode = codeAdapter.GetVbCodeInfo(e.FilePath).VbCode;
                mc.ChangeDocument(e.FilePath, vbCode);
                logger.Info("DocumentChanged");
            };
            server.CompletionReq += (object sender, CompletionEventArgs e) => {
                e.Items = new List<CompletionItem>();
				if (!codeAdapter.Has(e.FilePath)) {
                    logger.Info($"CompletionReq, non: {Path.GetFileName(e.FilePath)}");
                    return;
				}
               
                codeAdapter.parse(e.FilePath, e.Text, out VbCodeInfo vbCodeInfo);
                var vbCode = vbCodeInfo.VbCode;
                var posOffset = vbCodeInfo.PositionOffset;
                var line = e.Line - vbCodeInfo.LineOffset;
                if (line < 0) {
                    logger.Info($"CompletionReq, line={line}: {Path.GetFileName(e.FilePath)}");
                    return;
                }
                var Items = mc.GetCompletions(e.FilePath, vbCode, line, e.Chara).Result;
                e.Items = Items;
                logger.Info("CompletionReq");
            };
            server.DefinitionReq += (object sender, DefinitionEventArgs e) => {
                var list = new List<DefinitionItem>();
                if (!codeAdapter.Has(e.FilePath)) {
                    e.Items = list;
                    logger.Info($"DefinitionReq, non: {Path.GetFileName(e.FilePath)}");
                    return;
                }
                var vbCodeInfo = codeAdapter.GetVbCodeInfo(e.FilePath);
                var vbCode = vbCodeInfo.VbCode;
                var posOffset = vbCodeInfo.PositionOffset;
                var line = e.Line - vbCodeInfo.LineOffset;
                if (line < 0) {
                    e.Items = list;
                    logger.Info($"DefinitionReq, line={line}: {Path.GetFileName(e.FilePath)}");
                    return;
                }
                var Items = mc.GetDefinitions(e.FilePath, vbCode, line, e.Chara).Result;
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
                logger.Info("DefinitionReq");
            };
            server.HoverReq += (object sender, CompletionEventArgs e) => {
                var list = new List<CompletionItem>();
                if (!codeAdapter.Has(e.FilePath)) {
                    e.Items = list;
                    logger.Info($"HoverReq, non: {Path.GetFileName(e.FilePath)}");
                    return;
                }
                var vbCodeInfo = codeAdapter.GetVbCodeInfo(e.FilePath);
                var vbCode = vbCodeInfo.VbCode;
                var posOffset = vbCodeInfo.PositionOffset;
                var line = e.Line - vbCodeInfo.LineOffset;
                if (line < 0) {
                    e.Items = list;
                    logger.Info($"HoverReq, non: {Path.GetFileName(e.FilePath)}");
                    return;
                }
                var Items = mc.GetDefinitions(e.FilePath, vbCode, line, e.Chara).Result;
                foreach (var item in Items) {
                    var sp = item.Start.Positon;
                    var ep = item.End.Positon;
                    var hoverItem = mc.GetHover(item.FilePath, e.Text, (int)((sp + ep) / 2)).Result;
                    list.Add(hoverItem);
                }
                e.Items = list;
                logger.Info("HoverReq");
            };
            server.DiagnosticReq += (object sender, DiagnosticEventArgs e) => {
                if (!codeAdapter.Has(e.FilePath)) {
                    logger.Info($"DiagnosticReq, non: {Path.GetFileName(e.FilePath)}");
                    return;
                }
                var vbCodeInfo = codeAdapter.GetVbCodeInfo(e.FilePath);
                var lineOffset = vbCodeInfo.LineOffset;
                var items = mc.GetDiagnostics(e.FilePath).Result;
                foreach (var item in items) {
                    item.StartLine += lineOffset;
                    item.EndLine += lineOffset;
                }
                e.Items = items;
                logger.Info("DiagnosticReq");
            };
            server.DebugGetDocumentsEvent += (object sender, DebugEventArgs e) => {
                e.Text = JsonSerializer.Serialize(codeAdapter.getVbCodeDict());
            };
			server.ReferencesReq += (object sender, ReferencesEventArgs e) => {
                if (!codeAdapter.Has(e.FilePath)) {
                    logger.Info($"ReferencesReq, non: {Path.GetFileName(e.FilePath)}");
                    return;
                }
                var vbCodeInfo = codeAdapter.GetVbCodeInfo(e.FilePath);
                var line = e.Line - vbCodeInfo.LineOffset;
                var items = mc.GetReferences(e.FilePath, line, e.Chara).Result;
                foreach (var item in items) {
					if (codeAdapter.Has(item.FilePath)) {
                        var lineOffset = codeAdapter.GetVbCodeInfo(item.FilePath).LineOffset; 
                        item.Start.Line += lineOffset;
                        item.End.Line += lineOffset;
                    }
                }
                e.Items = items;
                logger.Info("ReferencesReq");
            };
            server.SignatureHelpReq += (object sender, SignatureHelpEventArgs e) => {
                var items = new List<SignatureHelpItem>();
                if (!codeAdapter.Has(e.FilePath)) {
                    e.Items = items;
                    logger.Info($"SignatureHelpReq, non: {Path.GetFileName(e.FilePath)}");
                    return;
                }
                codeAdapter.parse(e.FilePath, e.Text, out VbCodeInfo vbCodeInfo);
                var vbCode = vbCodeInfo.VbCode;
                var posOffset = vbCodeInfo.PositionOffset;
                var line = e.Line - vbCodeInfo.LineOffset;
                if (line < 0) {
                    e.Items = items;
                    logger.Info($"SignatureHelpReq, line < 0: {Path.GetFileName(e.FilePath)}");
                    return;
                }
                var (procLine, procCharaPos, argPosition) = mc.GetSignaturePosition(e.FilePath, vbCode, line, e.Chara);
                if (procLine < 0) {
                    e.Items = items;
                    logger.Info($"SignatureHelpReq, procLine < 0: {Path.GetFileName(e.FilePath)}");
                    return;
                }

                var Items = mc.GetDefinitions(e.FilePath, vbCode, procLine, procCharaPos).Result;
                foreach (var item in Items) {
                    var sp = item.Start.Positon;
                    var ep = item.End.Positon;
                    var sigItem = mc.GetSignatureHelp(item.FilePath, (int)((sp + ep) / 2)).Result;
                    if(sigItem != null) {
                        sigItem.ActiveParameter = argPosition;
                        items.Add(sigItem);
                    }
                }
                e.Items = items;
                logger.Info("SignatureHelpReq");
            };

            logger.Info("Initialized");
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
