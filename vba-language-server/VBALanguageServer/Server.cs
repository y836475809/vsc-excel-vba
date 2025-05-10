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
				try {
					var vbCode = this.vbaca.Rewrite(fp, Util.GetCode(fp));
					this.vbCache[fp] = vbCode;
					this.vbaca.AddDocument(fp, vbCode, true);
				} catch (Exception ex) {
#if DEBUG
					Logger.Error($"{Path.GetFileName(fp)}: {ex.Message}, {ex.StackTrace}");
#else
					Logger.Error($"{Path.GetFileName(fp)}: {ex.Message}");
#endif
				}
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
					Save = new LSP.SaveOptions {
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
			this.SendNotificationAsync("custom/initialized");
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
			var vbaDiagnostics = this.vbaca.GetDiagnostics(fp).Result;
			var diagnostics = vbaDiagnostics.Select(x => {
				return new LSP.Diagnostic {
					Code = x.ID,
					Severity = Util.ToSeverity(x.Severity),
					Message =x.Message,
					Range = new() {
						Start = new() { Line = x.Start.Item1, Character = x.Start.Item2 },
						End = new() { Line = x.End.Item1, Character = x.End.Item2 }
					}
				};
			});
			var diagnosticParams = new LSP.PublishDiagnosticParams {
				Uri = uri,
				Diagnostics = [..diagnostics]
			};
			this.SendNotificationAsync(Methods.TextDocumentPublishDiagnostics, diagnosticParams);
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

			var vbaHover = vbaca.GetHover(fp, line, chara).Result;
			if(vbaHover == null) {
				return null;
			}
			var contents = new List<LSP.SumType<string, LSP.MarkedString>>();
			foreach (var content in vbaHover.Contents){
				contents.Add(new LSP.MarkedString {
					Language = content.Language,
					Value = content.Value
				});
			}
			if (contents.Count == 0) {
				return null;
			}
			var hover = new Hover() {
				Contents = contents.ToArray(),
				Range = new LSP.Range() {
					Start = new(vbaHover.Start.Item1, vbaHover.Start.Item2),
					End = new(vbaHover.Start.Item1, vbaHover.Start.Item1),
				},
			};
			return hover;
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

			var vbaLocations = vbaca.GetReferences(fp, line, chara).Result;
			var locations = vbaLocations.Select(x => {
				return new LSP.Location {
					Uri = x.Uri,
					Range = new() {
						Start = new() { Line = x.Start.Item1, Character = x.Start.Item2 },
						End = new() { Line = x.End.Item1, Character = x.End.Item2 }
					},
				};
			});
			
			return [..locations];
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
			
			var vbaLocations = vbaca.GetDefinitions(fp, line, chara).Result;
			var locations = vbaLocations.Select(x => {
				return new LSP.Location {
					Uri = x.Uri,
					Range = new() {
						Start = new() { Line = x.Start.Item1, Character = x.Start.Item2 },
						End = new() { Line = x.End.Item1, Character = x.End.Item2 }
					},
				};
			});
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

			var items = vbaca.GetCompletions(fp, line, chara).Result;
			var completionItems = items.Select(item => {
				return new LSP.CompletionItem {
					Label = item.Label,
					Data = item.Display,
					Documentation = item.Doc,
					Kind = Util.ToKind(item.Kind)
				};
			});
			foreach (var item in items) {
				this.CompletionResolveDict[Convert.ToString(item.Display)] = Convert.ToString(item.Doc);
			}

			var completionList = new CompletionList() {
				IsIncomplete = false,
				Items = [.. completionItems],
			};
			return completionList;
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

			var (argPosition, items) = vbaca.GetSignatureHelp(fp, line, chara).Result;
			if(argPosition < 0) {
				return null;
			}

			var ret = new List<LSP.SignatureInformation>();
			foreach (var item in items) {
				var paramInfos = item.ParameterInfos.Select(x => {
					return new LSP.ParameterInformation {
						Label = x.Label,
						Documentation = x.Doc
					};
				});
				ret.Add(new LSP.SignatureInformation {
					Label = item.Label,
					Documentation = new LSP.MarkupContent {
						Kind = LSP.MarkupKind.Markdown,
						Value = $"```xml\n{item.Doc}\n```"
					},
					Parameters = [.. paramInfos]
				});
			}
			var signatureHelp = new LSP.SignatureHelp {
				ActiveParameter = argPosition,
				Signatures = [..ret]
			};
			if (signatureHelp.Signatures.Length == 0) {
				return null;
			}
			return signatureHelp;
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
			var vbaDocSymbols = this.vbaca.GetDocumentSymbols(fp, uri);
			var docSymbols = vbaDocSymbols.Select(vbaSymbol => {
				var start = vbaSymbol.Start;
				var end = vbaSymbol.End;
				var range = new LSP.Range {
					Start = new() { Line = start.Item1, Character = start.Item2 },
					End = new() { Line = end.Item1, Character = end.Item2 }
				};
				var docSymbol = new DocumentSymbol {
					Name = vbaSymbol.Name,
					Kind = Util.ToSymbolKind(vbaSymbol.Kind),
					Deprecated = false,
					Range = range,
					SelectionRange = range
				};
				var children = vbaSymbol.Children.Select(child => {
					var childStart = child.Start;
					var childEnd = child.End;
					var childRange = new LSP.Range {
						Start = new() { Line = childStart.Item1, Character = childStart.Item2 },
						End = new() { Line = childEnd.Item1, Character = childEnd.Item2 }
					};
					return new DocumentSymbol {
						Name = child.Name,
						Kind = Util.ToSymbolKind(child.Kind),
						Deprecated = false,
						Range = childRange,
						SelectionRange = childRange
					};
				});
				docSymbol.Children = [..children];
				return docSymbol;
			});	
			return [..docSymbols];
		}

		private Task SendNotificationAsync<TIn>(LspNotification<TIn> method, TIn param) {
			return this.rpc.NotifyWithParameterObjectAsync(method.Name, param);
		}

		private Task SendNotificationAsync(string method) {
			return this.rpc.NotifyWithParameterObjectAsync(method);
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
