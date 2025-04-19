using Microsoft.CodeAnalysis;
using Microsoft.VisualStudio.LanguageServer.Protocol;
using LSP = Microsoft.VisualStudio.LanguageServer.Protocol;
using Nerdbank.Streams;
using Newtonsoft.Json.Linq;
using StreamJsonRpc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using VBACodeAnalysis;
using System.Diagnostics;
using System.Reflection.Emit;
using static System.Formats.Asn1.AsnWriter;
using SharpCompress.Common;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;


namespace VBALanguageServer {
	internal class CurrentTextDocument {
		public string AbsoluteUri;
		public string Text;

		public void Update(DidChangeTextDocumentParams @params) {
			var changes = @params.ContentChanges;
			if (changes.Length == 0) {
				return;
			}
			this.AbsoluteUri = @params.TextDocument.Uri.AbsoluteUri;
			this.Text = changes[0].Text;
		}

		public bool TryGetText(Uri uri, out string text) {
			text = null;
			if (this.AbsoluteUri != uri.AbsoluteUri) {
				return false;
			}
			text = this.Text;
			return true;
		}
	}

	public class Server {
		private readonly string srcDirName;
		private VBACodeAnalysis.VBACodeAnalysis vbaca;
		private Dictionary<string, string> vbCache;
		private readonly JsonRpc rpc;
		private DebounceDispatcher didTextChangeDebounce;
		private CurrentTextDocument currentTextDocument;

		public Server(string srcDirName, Stream stdin, Stream stdout) {
			this.srcDirName = srcDirName;
			var stream = FullDuplexStream.Splice(stdin, stdout);
			this.rpc = JsonRpc.Attach(stream, this);
			this.rpc.Completion.Wait();
		}

		private void InitVBACodeAnalysis(string work_dir, string[] exts) {
			Logger.Info($"work_dir={work_dir}");

			this.didTextChangeDebounce = new DebounceDispatcher(300);
			this.currentTextDocument = new CurrentTextDocument();

			this.vbaca = new VBACodeAnalysis.VBACodeAnalysis();
			this.vbaca.setSetting(this.LoadSettings().RewriteSetting);
			this.vbCache = [];

			var fps = Enumerable.Empty<string>();
			foreach (var ext in exts) {
				fps = fps.Concat(Directory.GetFiles(work_dir, ext));
			}
			fps = fps.Concat(GetVBDefineFiles());
			Logger.Info($"fps={fps.Count()}");
			foreach (var fp in fps) {
				var vbCode = this.vbaca.Rewrite(fp, Util.GetCode(fp));
				this.vbCache[fp] = vbCode;
				this.vbaca.AddDocument(fp, vbCode, true);
			}
			//this.vbaca.ApplyChanges([.. fps]);
		}

		[JsonRpcMethod(Methods.InitializeName)]
		public object Initialize(JToken arg) {
			Logger.Info("Initialize");

			var init_params = arg.ToObject<InitializeParams>();
			var work_dir = Path.Combine(this.GetFsPath(init_params.RootUri), this.srcDirName);
			this.InitVBACodeAnalysis(work_dir, ["*.bas", "*.cls"]);

			var capabilities = new ServerCapabilities {
				TextDocumentSync = new TextDocumentSyncOptions {
					OpenClose = true,
					Change = TextDocumentSyncKind.Full,
					Save = new SaveOptions {
						IncludeText = false,
					}
				},
				CompletionProvider = new CompletionOptions {
					TriggerCharacters = ["."],
					ResolveProvider = true,
					WorkDoneProgress = false,
				},
				HoverProvider = true,
				SignatureHelpProvider = new SignatureHelpOptions {
					TriggerCharacters = ["(", ","],
				},
				DefinitionProvider = true,
				TypeDefinitionProvider = false,
				ImplementationProvider = false,
				ReferencesProvider = true,
				DocumentHighlightProvider = false,
				DocumentSymbolProvider = true,
				CodeLensProvider = null,
				DocumentLinkProvider = null,
				DocumentFormattingProvider = false,
				DocumentRangeFormattingProvider = false,
				RenameProvider = false,
				FoldingRangeProvider = null,
				ExecuteCommandProvider = null,
				WorkspaceSymbolProvider = false,
			};

			var result = new InitializeResult {
				Capabilities = capabilities
			};
			return result;
		}

		[JsonRpcMethod(Methods.InitializedName)]
		public void Initialized(JToken arg) {
			Logger.Info("Initialized");
		}

		[JsonRpcMethod(Methods.ShutdownName)]
		public JToken ShutdownName() {
			return null;
		}

		[JsonRpcMethod(Methods.ExitName)]
		public void ExitName() {
			System.Environment.Exit(0);
		}

		[JsonRpcMethod(Methods.WorkspaceDidChangeWatchedFilesName)]
		public ApplyWorkspaceEditResponse OnWorkspaceDidChangeWatchedFiles(JToken arg) {
			Logger.Info("WorkspaceDidChangeWatchedFiles");
			var @params = arg.ToObject<DidChangeWatchedFilesParams>();
			var createFPList = @params.Changes
				.Where(x => x.FileChangeType == FileChangeType.Created)
				.Select(x => {
					var fp = this.GetFsPath(x.Uri);
					// TODO
					// [42:32.074][Info] WorkspaceDidChangeWatchedFiles
					// [42:32.075][Info] Created \test-data\src\m2.bas
					// [42:32.096][Info] Created \test-data\src\m2.bas
					Logger.Info($"Created {fp}");
					var vbCode = vbaca.Rewrite(fp, Util.GetCode(fp));
					this.vbCache[fp] = vbCode;
					this.vbaca.AddDocument(fp, vbCode, false);
					return fp;
				});
			if(createFPList.Any()) {
				this.vbaca.ApplyChanges([.. createFPList]);
			}

			@params.Changes
				.Where(x => x.FileChangeType == FileChangeType.Deleted)
				.ToList().ForEach(x => {
					var fp = this.GetFsPath(x.Uri);
					Logger.Info($"Deleted {fp}");
					this.vbaca.DeleteDocument(fp);
					this.vbCache.Remove(fp);
				});
			var result = new ApplyWorkspaceEditResponse {
				Applied = true,
			};
			return result;
		}

		[JsonRpcMethod(Methods.TextDocumentDidOpenName)]
		public void OnTextDocumentOpened(JToken arg) {
			Logger.Info("OnTextDocumentOpened");
			var @params = arg.ToObject<DidOpenTextDocumentParams>();
			this.SendDiagnostics(@params.TextDocument.Uri);
		}

		[JsonRpcMethod(Methods.TextDocumentDidCloseName)]
		public void OnTextDocumentClosed(JToken arg) {
			var @params = arg.ToObject<DidCloseTextDocumentParams>();
			var diag_params = new PublishDiagnosticParams {
				Uri = @params.TextDocument.Uri,
				Diagnostics = []
			};
			this.SendNotificationAsync(Methods.TextDocumentPublishDiagnostics, diag_params);
		}

		[JsonRpcMethod(Methods.TextDocumentDidChangeName)]
		public void OnTextDocumentChanged(JToken arg) {
			var @params = arg.ToObject<DidChangeTextDocumentParams>();
			this.currentTextDocument.Update(@params);

			this.didTextChangeDebounce.Debounce(() => {
				Logger.Info("OnTextDocumentChanged");
				var @params = arg.ToObject<DidChangeTextDocumentParams>();
				var uri = @params.TextDocument.Uri;
				var changes = @params.ContentChanges;
				Logger.Info(uri.LocalPath);
				var fp = this.GetFsPath(uri);
				if (changes.Length == 0) {
					return;
				}
				var vbCode = this.vbaca.Rewrite(fp, changes[0].Text);
				this.vbCache[fp] = vbCode;
				this.vbaca.ChangeDocument(fp, vbCode);
				this.SendDiagnostics(@params.TextDocument.Uri);
			});
		}

		public void SendDiagnostics(Uri uri) {
			var fp = this.GetFsPath(uri);
			var items = this.vbaca.GetDiagnostics(fp).Result;
			var parameter = new PublishDiagnosticParams();
			var diagnostics = new List<LSP.Diagnostic>();
			foreach (var item in items) {
				var charDiff = this.vbaca.GetCharaDiff(fp, item.StartLine, item.StartChara);
				item.StartChara -= charDiff;
				item.EndChara -= charDiff;

				var diagnostic = new LSP.Diagnostic();
				diagnostic.Message = item.Message;
				var severity = LSP.DiagnosticSeverity.Information;
				if (item.Severity.ToLower() == "info") {
					severity = LSP.DiagnosticSeverity.Information;
				}
				if (item.Severity.ToLower() == "warning") {
					severity = LSP.DiagnosticSeverity.Warning;
				}
				if (item.Severity.ToLower() == "error") {
					severity = LSP.DiagnosticSeverity.Error;
				}
				diagnostic.Severity = severity;
				diagnostic.Range = new LSP.Range {
					Start = new Position(item.StartLine, item.StartChara),
					End = new Position(item.EndLine, item.EndChara)
				};
				diagnostics.Add(diagnostic);
			}

			parameter.Uri = uri;
			parameter.Diagnostics = [.. diagnostics];
			this.SendNotificationAsync(Methods.TextDocumentPublishDiagnostics, parameter);
		}

		[JsonRpcMethod(Methods.TextDocumentHoverName)]
		public Hover OnTextDocumentHover(JToken arg) {
			var @params = arg.ToObject<TextDocumentPositionParams>();
			var fp = this.GetFsPath(@params.TextDocument.Uri);
			if (!this.vbCache.ContainsKey(fp)) {
				Logger.Info($"{Path.GetFileName(fp)}");
				return null;
			}
			var line = @params.Position.Line;
			var chara = @params.Position.Character;
			if (line < 0) {
				Logger.Info($"HoverReq, non: {Path.GetFileName(fp)}");
				return null;
			}
			var adjChara = vbaca.GetCharaDiff(fp, line, chara) + chara;
			var items = vbaca.GetHover(fp, line, adjChara).Result;
			if(items.Count == 0) {
				return null;
			}
			var item = items[0];
			var msList = new List<SumType<string, MarkedString>>();
			if (item.Description != "") {
				msList.Add(new MarkedString{
					Language = "xml",
					Value = item.Description
				});
			}
			if (item.DisplayText != "") {
				msList.Add(new MarkedString {
					Language = "vb",
					Value = item.DisplayText
				});
			}
			if (item.ReturnType != "") {
				msList.Add(new MarkedString {
					Language = MarkupKind.PlainText.ToString(),
					Value = $"@return {item.ReturnType}"
				});
			}
			if (item.Kind != "") {
				msList.Add(new MarkedString {
					Language = MarkupKind.PlainText.ToString(),
					Value = $"@kind {item.Kind}"
				});
			}
			var result = new Hover() {
				Contents = msList.ToArray(),
				Range = new LSP.Range() {
					Start = new Position(line, adjChara),
					End = new Position(line, adjChara),
				},
			};
			return result;
		}

		[JsonRpcMethod(Methods.TextDocumentReferencesName)]
		public LSP.Location[] OnTextDocumentReferences(JToken arg) {
			Logger.Info("OnTextDocumentReferences");
			var @params = arg.ToObject<ReferenceParams>();
			var fp = this.GetFsPath(@params.TextDocument.Uri);
			if (!this.vbCache.ContainsKey(fp)) {
				Logger.Info($"ReferencesReq, non: {Path.GetFileName(fp)}");
				return [];
			}
			var line = @params.Position.Line;
			var chara = @params.Position.Character;
			if (line < 0) {
				Logger.Info($"DefinitionReq, line={line}: {Path.GetFileName(fp)}");
				return [];
			}
			var locations = new List<LSP.Location>();
			var adjChara = vbaca.GetCharaDiff(fp, line, chara) + chara;
			var items = vbaca.GetReferences(fp, line, adjChara).Result;
			foreach (var item in items) {
				var start = AdjustPosition(item.FilePath, item.Start.Line, item.Start.Character);
				var end = AdjustPosition(item.FilePath, item.End.Line, item.End.Character);
				var loc = new LSP.Location {
					Uri = new Uri(item.FilePath),
					Range = new LSP.Range {
						Start = start,
						End = end,
					}
				};
				locations.Add(loc);
			}
			return locations.ToArray();
		}

		[JsonRpcMethod(Methods.TextDocumentDefinitionName)]
		public LSP.Location[] OnTextDocumentDefinition(JToken arg) {
			Logger.Info("OnTextDocumentDefinition");
			var @params = arg.ToObject<TextDocumentPositionParams>();
			var fp = this.GetFsPath(@params.TextDocument.Uri);
			if (!this.vbCache.TryGetValue(fp, out string vbCode)) {
				Logger.Info($"DefinitionReq, non: {Path.GetFileName(fp)}");
				return [];
			}

			var line = @params.Position.Line;
			var chara = @params.Position.Character;
			if (line < 0) {
				Logger.Info($"DefinitionReq, line={line}: {Path.GetFileName(fp)}");
				return [];
			}
			var locations = new List<LSP.Location>();
			var adjChara = vbaca.GetCharaDiff(fp, line, chara) + chara;
			var Items = vbaca.GetDefinitions(fp, vbCode, line, adjChara).Result;
			foreach (var item in Items) {
				if (item.IsKindClass()) {
					var loc = new LSP.Location {
						Uri = new Uri(item.FilePath),
						Range = new LSP.Range {
							Start = new Position(0, 0),
							End = new Position(0, 0),
						}
					};
					locations.Add(loc);
				} else {
					var start = AdjustPosition(item.FilePath, item.Start.Line, item.Start.Character);
					var end = AdjustPosition(item.FilePath, item.End.Line, item.End.Character);
					var loc = new LSP.Location {
						Uri = new Uri(item.FilePath),
						Range = new LSP.Range {
							Start = start,
							End = end,
						}
					};
					locations.Add(loc);
				}
			}
			return [.. locations];
		}

		private Dictionary<string, string> CompletionResolveDict = [];

		[JsonRpcMethod(Methods.TextDocumentCompletionName)]
		public CompletionList OnTextDocumentCompletion(JToken arg) {
			Logger.Info("OnTextDocumentCompletion");
			var @params = arg.ToObject<CompletionParams>();
			this.CompletionResolveDict = [];

			var fp = this.GetFsPath(@params.TextDocument.Uri);
			var line = @params.Position.Line;
			var chara = @params.Position.Character;
			if (line < 0) {
				Logger.Info($"CompletionReq, line={line}: {Path.GetFileName(fp)}");
				return null;
			}

			if (!this.vbCache.TryGetValue(fp, out string vbCode)) {
				Logger.Info($"CompletionReq, non: {Path.GetFileName(fp)}");
				return null;
			}
			if (this.currentTextDocument.TryGetText(@params.TextDocument.Uri, out string currentText)) {
				if (currentText != vbCode) {
					var currentvVBCode = vbaca.Rewrite(fp, currentText);
					vbaca.ChangeDocument(fp, currentvVBCode);
				}
			}

			List<LSP.CompletionItem> compItems = new List<LSP.CompletionItem>();
			var adjChara = vbaca.GetCharaDiff(fp, line, chara) + chara;
			var items = vbaca.GetCompletions(fp, vbCode, line, adjChara).Result;
			foreach (var item in items) {
				var compItem = new LSP.CompletionItem();
				compItem.Data = item.CompletionText;
				this.CompletionResolveDict[item.CompletionText] = item.Description;

				compItem.Label = item.DisplayText;
				var kind = item.Kind.ToLower();
				var compItemKind = CompletionItemKind.Text;
				if (kind == "method") { compItemKind = CompletionItemKind.Method; }
				if (kind == "field") { compItemKind = CompletionItemKind.Field; }
				if (kind == "property") { compItemKind = CompletionItemKind.Property; }
				if (kind == "local") { compItemKind = CompletionItemKind.Variable; }
				if (kind == "class") { compItemKind = CompletionItemKind.Class; }
				if (kind == "keyword") { compItemKind = CompletionItemKind.Keyword; }
				compItem.Kind = compItemKind;
				compItems.Add(compItem);
			}
			var list = new CompletionList() {
				IsIncomplete =false,
				Items = [.. compItems],
			};
			return list;
		}

		[JsonRpcMethod(Methods.TextDocumentCompletionResolveName)]
		public LSP.CompletionItem OnTextDocumentCompletionResolve(JToken arg) {
			Logger.Info("OnTextDocumentCompletionResolve");
			var @params = arg.ToObject<LSP.CompletionItem>();

			var completionText = Convert.ToString(@params.Data);
			var item = new LSP.CompletionItem();
			var value = $"```vb\n{completionText}\n \n```";
			if (this.CompletionResolveDict.TryGetValue(completionText, out string doc)) {
				value = $"```vb\n{completionText}\n \n```\n```xml\n{doc}\n```";
			}
			item.Documentation = new MarkupContent {
				Kind = MarkupKind.Markdown,
				Value = value,
			};
			return item;
		}

		[JsonRpcMethod(Methods.TextDocumentSignatureHelpName)]
		public SignatureHelp OnTextDocumentSignatureHelp(JToken arg) {
			Logger.Info("OnTextDocumentSignatureHelp");
			var @params = arg.ToObject<SignatureHelpParams>();

			var fp = this.GetFsPath(@params.TextDocument.Uri);
			var line = @params.Position.Line;
			var chara = @params.Position.Character;
			if (line < 0) {
				Logger.Info($"CompletionReq, line={line}: {Path.GetFileName(fp)}");
				return null;
			}
			
			if (!this.vbCache.TryGetValue(fp, out string vbCode)) {
				return null;
			}
			if (this.currentTextDocument.TryGetText(@params.TextDocument.Uri, out string currentText)) {
				if (currentText != vbCode) {
					var currentvVBCode = vbaca.Rewrite(fp, currentText);
					vbaca.ChangeDocument(fp, currentvVBCode);
				}
			}

			var adjChara = vbaca.GetCharaDiff(fp, line, chara) + chara;
			var (procLine, procCharaPos, argPosition) = vbaca.GetSignaturePosition(fp, line, adjChara);
			if (procLine < 0) {
				Logger.Info($"Sigine < 0: {Path.GetFileName(fp)}");
				return null;
			}

			var items = vbaca.GetSignatureHelp(fp, procLine, procCharaPos).Result;
			if (!items.Any()) {
				Logger.Info($"items == 0: {Path.GetFileName(fp)}");
				return null;
			}

			var signatures = new List<SignatureInformation>();
			foreach (var item in items) {
				signatures.Add(new SignatureInformation {
					Label = item.DisplayText,
					Documentation = new MarkupContent {
						Kind = MarkupKind.PlainText,
						Value = item.Description,
					},
					Parameters = item.Args.Select(args => {
						return new ParameterInformation {
							Label = args.Name,
							Documentation = args.AsType
						};
					}).ToArray()
				});
			}
			var resultl = new SignatureHelp {
				ActiveParameter = argPosition,
				Signatures = [..signatures]
			};
			return resultl;
		}

		[JsonRpcMethod(Methods.TextDocumentDocumentSymbolName)]
		public DocumentSymbol[] OnTextDocumentDocumentSymbol(JToken arg) {
			while (!didTextChangeDebounce.IsCompleted) {
				Task.Delay(100).Wait();
			}
			Logger.Info("OnTextDocumentDocumentSymbol");
			var @params = arg.ToObject<DocumentSymbolParams>();
			var uri = @params.TextDocument.Uri;
			var fp = this.GetFsPath(@params.TextDocument.Uri);
			var docSymbols = this.vbaca.GetDocumentSymbols(fp, uri);
			return docSymbols;
		}

		private Task SendNotificationAsync<TIn>(LspNotification<TIn> method, TIn param) {
			return this.rpc.NotifyWithParameterObjectAsync(method.Name, param);
		}

		private Position AdjustPosition(string filePath, int vbaLine, int vbaChara) {
			var charaDiff = vbaca.GetCharaDiff(filePath, vbaLine, vbaChara);
			var line = vbaLine;
			var chara = vbaChara - charaDiff;
			if (line < 0) {
				line = 0;
			}
			if (chara < 0) {
				chara = 0;
			}
			return new LSP.Position(line, chara);
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

		private List<string> GetVBDefineFiles() {
			var assembly = Assembly.GetEntryAssembly();
			var dir = Path.Join(Path.GetDirectoryName(assembly.Location), "d.vb");
			if (!Path.Exists(dir)) {
				return [];
			}
			var fps = Directory.GetFiles(dir, "*.d.vb");
			return [.. fps];
		}

		private string GetFsPath(Uri uri) {
			string lp = uri.LocalPath.TrimStart('/');
			string fsPath = Path.GetFullPath(lp);
			return fsPath;
		}
	}
}
