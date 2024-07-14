﻿using VBACodeAnalysis;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Reflection;
using System.Linq;

namespace VBALanguageServer {
	class App {
        private Server server;
		private VBACodeAnalysis.VBACodeAnalysis vbaca;
        private Logger logger;

		private PreprocVBA _preprocVba;
		private Dictionary<string, string> _vbCache;

		public App() {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            logger = new Logger();
        }

        public void Run(int port) {
            server.Setup(port);
            server.Run();
        }

        private void Reset() {
            _preprocVba = new PreprocVBA();
            _vbCache = [];

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
                    var vbCode = _preprocVba.Rewrite(FilePath, Helper.getCode(FilePath));
                    _vbCache[FilePath] = vbCode;
                    vbaca.AddDocument(FilePath, vbCode, false);
                }
                vbaca.ApplyChanges(e.FilePaths);
                logger.Info("DocumentAdded");
            };
            server.DocumentDeleted += (object sender, DocumentDeletedEventArgs e) => {
                foreach (var FilePath in e.FilePaths) {
                    vbaca.DeleteDocument(FilePath);
                    _vbCache.Remove(FilePath);
                }
				logger.Info("DocumentDeleted");
            };
            server.DocumentRenamed += (object sender, DocumentRenamedEventArgs e) => {
                vbaca.DeleteDocument(e.OldFilePath);
				var filePath = e.NewFilePath;
                var vbCode = _preprocVba.Rewrite(filePath, Helper.getCode(filePath));
                _vbCache.Remove(filePath);
                _vbCache[filePath] = vbCode;
                vbaca.AddDocument(e.NewFilePath, vbCode);
                logger.Info("DocumentRenamed");
            };
            server.DocumentChanged += (object sender, DocumentChangedEventArgs e) => {
				var filePath = e.FilePath;
                var vbCode = _preprocVba.Rewrite(filePath, e.Text);
                _vbCache[filePath] = vbCode;
                vbaca.ChangeDocument(e.FilePath, vbCode);
                logger.Info("DocumentChanged");
            };
            server.CompletionReq += (object sender, CompletionEventArgs e) => {
                e.Items = new List<CompletionItem>();
                if (!_vbCache.ContainsKey(e.FilePath)) {
                    logger.Info($"CompletionReq, non: {Path.GetFileName(e.FilePath)}");
                    return;
				}
				var filePath = e.FilePath;
                var vbCode = _preprocVba.Rewrite(filePath, e.Text);
                var line = e.Line;
                if (line < 0) {
                    logger.Info($"CompletionReq, line={line}: {Path.GetFileName(e.FilePath)}");
                    return;
                }

                vbaca.ChangeDocument(e.FilePath, vbCode);
                var adjChara = vbaca.GetCharaDiff(e.FilePath, line, e.Chara) + e.Chara;
                var Items = vbaca.GetCompletions(e.FilePath, vbCode, line, adjChara).Result;
                e.Items = Items;
                logger.Info("CompletionReq");
            };
            server.DefinitionReq += (object sender, DefinitionEventArgs e) => {
                var list = new List<DefinitionItem>();
                var filePath = e.FilePath;
                if (!_vbCache.ContainsKey(filePath)) {
                    e.Items = list;
                    logger.Info($"DefinitionReq, non: {Path.GetFileName(e.FilePath)}");
                    return;
                }
                var vbCode = _vbCache[filePath];
                var line = e.Line;
                if (line < 0) {
                    e.Items = list;
                    logger.Info($"DefinitionReq, line={line}: {Path.GetFileName(filePath)}");
                    return;
                }
                var adjChara = vbaca.GetCharaDiff(e.FilePath, line, e.Chara) + e.Chara;
                var Items = vbaca.GetDefinitions(e.FilePath, vbCode, line, adjChara).Result;
                foreach (var item in Items) {
                    if (item.IsKindClass()) {
                        item.Start.Positon = 0;
                        item.Start.Line = 0;
                        item.Start.Character = 0;
                        item.End.Positon = 0;
                        item.End.Line = 0;
                        item.End.Character = 0;
                    } else {
                        item.Start = AdjustLocation(item.FilePath,  item.Start);
                        item.End = AdjustLocation(item.FilePath, item.End);
                    }
                    list.Add(item);
                }
                e.Items = list;
                logger.Info("DefinitionReq");
            };
            server.HoverReq += (object sender, CompletionEventArgs e) => {
                e.Items = new List<CompletionItem>();
				var filePath = e.FilePath;
                if (!_vbCache.ContainsKey(filePath)) {
                    logger.Info($"HoverReq, non: {Path.GetFileName(e.FilePath)}");
                    return;
                }
                var vbCode = _vbCache[filePath];
                var line = e.Line;
                if (line < 0) {
                    logger.Info($"HoverReq, non: {Path.GetFileName(e.FilePath)}");
                    return;
                }
                var adjChara = vbaca.GetCharaDiff(e.FilePath, line, e.Chara) + e.Chara;
                e.Items = vbaca.GetHover(e.FilePath, line, adjChara).Result;
                logger.Info("HoverReq");
            };
            server.DiagnosticReq += (object sender, DiagnosticEventArgs e) => {
				var filePath = e.FilePath;
                if (!_vbCache.ContainsKey(filePath)) {
                    logger.Info($"DiagnosticReq, non: {Path.GetFileName(e.FilePath)}");
                    return;
                }
                var items1 = _preprocVba.GetDiagnostics(e.FilePath);
				var items2 = vbaca.GetDiagnostics(e.FilePath).Result;
                var items = items2.Concat(items1).ToList();
				foreach (var item in items) {
                    var charDiff = vbaca.GetCharaDiff(e.FilePath, item.StartLine, item.StartChara);
                    item.StartChara -= charDiff;
                    item.EndChara -= charDiff;
                }
                e.Items = items;
                logger.Info("DiagnosticReq");
            };
            server.DebugGetDocumentsEvent += (object sender, DebugEventArgs e) => {
                e.Text = JsonSerializer.Serialize(_vbCache);
            };
			server.ReferencesReq += (object sender, ReferencesEventArgs e) => {
				var filePath = e.FilePath;
                if (!_vbCache.ContainsKey(filePath)) {
                    logger.Info($"ReferencesReq, non: {Path.GetFileName(e.FilePath)}");
                    return;
                }
                var adjChara = vbaca.GetCharaDiff(e.FilePath, e.Line, e.Chara) + e.Chara;
                var items = vbaca.GetReferences(e.FilePath, e.Line, adjChara).Result;
                foreach (var item in items) {
                    item.Start = AdjustLocation(item.FilePath, item.Start);
                    item.End = AdjustLocation(item.FilePath, item.End);
                }
                e.Items = items;
                logger.Info("ReferencesReq");
            };
            server.SignatureHelpReq += (object sender, SignatureHelpEventArgs e) => {
                var items = new List<SignatureHelpItem>();
				var filePath = e.FilePath;
                if (!_vbCache.ContainsKey(filePath)) {
                    e.Items = items;
                    logger.Info($"SignatureHelpReq, non: {Path.GetFileName(e.FilePath)}");
                    return;
                }
                var vbCode = _preprocVba.Rewrite(filePath, e.Text);
                var line = e.Line;
                if (line < 0) {
                    e.Items = items;
                    logger.Info($"SignatureHelpReq, line < 0: {Path.GetFileName(e.FilePath)}");
                    return;
                }

                vbaca.ChangeDocument(e.FilePath, vbCode);
                var adjChara = vbaca.GetCharaDiff(e.FilePath, line, e.Chara) + e.Chara;
                var (procLine, procCharaPos, argPosition) = vbaca.GetSignaturePosition(e.FilePath, line, adjChara);
                if (procLine < 0) {
                    e.Items = items;
                    logger.Info($"SignatureHelpReq, procLine < 0: {Path.GetFileName(e.FilePath)}");
                    return;
                }

                items = vbaca.GetSignatureHelp(e.FilePath, procLine, procCharaPos).Result;
                foreach (var item in items) {
                    item.ActiveParameter = argPosition;
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

        private Location AdjustLocation(string defFilePath, Location location) {
            var charaDiff = vbaca.GetCharaDiff(defFilePath, location.Line, location.Character);
            location.Character -= charaDiff;
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
