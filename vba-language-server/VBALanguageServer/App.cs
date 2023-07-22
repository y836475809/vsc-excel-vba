﻿using VBACodeAnalysis;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Reflection;

namespace VBALanguageServer {
	class App {
		private CodeAdapter codeAdapter;
		private Server server;
		private VBACodeAnalysis.VBACodeAnalysis vbaca;
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
			vbaca = new VBACodeAnalysis.VBACodeAnalysis();
            var settings = LoadSettings();
            vbaca.setSetting(settings.RewriteSetting);
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
                    vbaca.AddDocument(FilePath, vbCode, false);
                }
                vbaca.ApplyChanges(e.FilePaths);
                logger.Info("DocumentAdded");
            };
            server.DocumentDeleted += (object sender, DocumentDeletedEventArgs e) => {
                foreach (var FilePath in e.FilePaths) {
                    vbaca.DeleteDocument(FilePath);
                    codeAdapter.Delete(FilePath);
                }
                logger.Info("DocumentDeleted");
            };
            server.DocumentRenamed += (object sender, DocumentRenamedEventArgs e) => {
                vbaca.DeleteDocument(e.OldFilePath);
                codeAdapter.Delete(e.OldFilePath);
                codeAdapter.SetCode(e.NewFilePath, Helper.getCode(e.NewFilePath));
                var vbCode = codeAdapter.GetVbCodeInfo(e.NewFilePath).VbCode;
                vbaca.AddDocument(e.NewFilePath, vbCode);
                logger.Info("DocumentRenamed");
            };
            server.DocumentChanged += (object sender, DocumentChangedEventArgs e) => {
                codeAdapter.SetCode(e.FilePath, e.Text);
                var vbCode = codeAdapter.GetVbCodeInfo(e.FilePath).VbCode;
                vbaca.ChangeDocument(e.FilePath, vbCode);
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
                var line = e.Line;
                if (line < 0) {
                    logger.Info($"CompletionReq, line={line}: {Path.GetFileName(e.FilePath)}");
                    return;
                }

                vbaca.ChangeDocument(e.FilePath, vbCode);
                var adjChara = vbaca.getoffset(e.FilePath, line, e.Chara) + e.Chara;
                var Items = vbaca.GetCompletions(e.FilePath, vbCode, line, adjChara).Result;
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
                var line = e.Line;
                if (line < 0) {
                    e.Items = list;
                    logger.Info($"DefinitionReq, line={line}: {Path.GetFileName(e.FilePath)}");
                    return;
                }
                var adjChara = vbaca.getoffset(e.FilePath, line, e.Chara) + e.Chara;
                var Items = vbaca.GetDefinitions(e.FilePath, vbCode, line, adjChara).Result;
                foreach (var item in Items) {
                    var itemVbCodeInfo = codeAdapter.GetVbCodeInfo(item.FilePath);
                    if (item.IsKindClass()) {
                        item.Start.Positon = 0;
                        item.Start.Line = 0;
                        item.Start.Character = 0;
                        item.End.Positon = 0;
                        item.End.Line = 0;
                        item.End.Character = 0;
                    } else {
                        item.Start = adjustLocation(item.FilePath,  item.Start);
                        item.End = adjustLocation(item.FilePath, item.End);
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
                var line = e.Line;
                if (line < 0) {
                    e.Items = list;
                    logger.Info($"HoverReq, non: {Path.GetFileName(e.FilePath)}");
                    return;
                }
                var adjChara = vbaca.getoffset(e.FilePath, line, e.Chara) + e.Chara;
                var Items = vbaca.GetDefinitions(e.FilePath, vbCode, line, adjChara).Result;
                foreach (var item in Items) {
                    var sp = item.Start.Positon;
                    var ep = item.End.Positon;
                    var hoverItem = vbaca.GetHover(item.FilePath, e.Text, (int)((sp + ep) / 2)).Result;
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
                var items = vbaca.GetDiagnostics(e.FilePath).Result;
                foreach (var item in items) {
                    var itemVbCodeInfo = codeAdapter.GetVbCodeInfo(e.FilePath);
                    int toVBAoffset = vbaca.getoffset(e.FilePath, item.StartLine, item.StartChara);
                    item.StartChara -= toVBAoffset;
                    item.EndChara -= toVBAoffset;
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
                var adjChara = vbaca.getoffset(e.FilePath, e.Line, e.Chara) + e.Chara;
                var items = vbaca.GetReferences(e.FilePath, e.Line, adjChara).Result;
                foreach (var item in items) {
                    var itemVbCodeInfo = codeAdapter.GetVbCodeInfo(item.FilePath);
                    item.Start = adjustLocation(item.FilePath, item.Start);
                    item.End = adjustLocation(item.FilePath, item.End);
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
                var line = e.Line;
                if (line < 0) {
                    e.Items = items;
                    logger.Info($"SignatureHelpReq, line < 0: {Path.GetFileName(e.FilePath)}");
                    return;
                }

                vbaca.ChangeDocument(e.FilePath, vbCode);
                var adjChara = vbaca.getoffset(e.FilePath, line, e.Chara) + e.Chara;
                var (procLine, procCharaPos, argPosition) = vbaca.GetSignaturePosition(e.FilePath, vbCode, line, adjChara);
                if (procLine < 0) {
                    e.Items = items;
                    logger.Info($"SignatureHelpReq, procLine < 0: {Path.GetFileName(e.FilePath)}");
                    return;
                }

                var Items = vbaca.GetDefinitions(e.FilePath, vbCode, procLine, procCharaPos).Result;
                foreach (var item in Items) {
                    var sp = item.Start.Positon;
                    var ep = item.End.Positon;
                    var sigItems = vbaca.GetSignatureHelp(item.FilePath, (int)((sp + ep) / 2)).Result;
                    foreach (var sigItem in sigItems) {
                        sigItem.ActiveParameter = argPosition;
                        items.Add(sigItem);
                    }
                }
                e.Items = items;
                logger.Info("SignatureHelpReq");
            };

            logger.Info("Initialized");
        }

        private Settings LoadSettings() {
            var settings = new Settings();
            var assembly = Assembly.GetEntryAssembly();
            var jsonPath = Path.Join(Path.GetDirectoryName(assembly.Location), "settings.json");
            using (var sr = new StreamReader(jsonPath)) {
                var jsonStr = sr.ReadToEnd();
                settings.Parse(jsonStr);
            }
            return settings;
        }

        private Location adjustLocation(string defFilePath, Location location) {
            int toVBAoffset = vbaca.getoffset(defFilePath, location.Line, location.Character);
            location.Character -= toVBAoffset;
            if (location.Line < 0) {
                location.Line = 0;
            }
            if (location.Positon < 0) {
                location.Positon = 0;
            }
            return new Location(location.Positon, location.Line, location.Character);
        }
    }
}
