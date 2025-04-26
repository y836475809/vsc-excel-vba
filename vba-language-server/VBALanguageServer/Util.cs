using System.IO;
using System.Text;
using System.Text.Json;
using LSP = Microsoft.VisualStudio.LanguageServer.Protocol;

namespace VBALanguageServer {
    class Util {
        public static string GetCode(string filePath) {
            var enc = Encoding.GetEncoding("shift_jis");
            using (var sr = new StreamReader(filePath, enc)) {
                return sr.ReadToEnd();
            }
        }

        public static LSP.CompletionItemKind ToKind(string kind) {
			switch (kind) {
                case "Method":
                    return LSP.CompletionItemKind.Method;
                case "Field":
					return LSP.CompletionItemKind.Field;
                case "Property":
					return LSP.CompletionItemKind.Property;
				case "Local":
					return LSP.CompletionItemKind.Variable;
				case "Class":
					return LSP.CompletionItemKind.Class;
				case "Keyword":
                    return LSP.CompletionItemKind.Keyword;
                default:
                    return LSP.CompletionItemKind.Keyword;
			}
        }

		public static LSP.DiagnosticSeverity ToSeverity(string kind) {
			switch (kind) {
				case "Info":
					return LSP.DiagnosticSeverity.Information;
				case "Warning":
					return LSP.DiagnosticSeverity.Warning;
				case "Error":
					return LSP.DiagnosticSeverity.Error;
				default:
					return LSP.DiagnosticSeverity.Information;
			}
		}

		public static LSP.SymbolKind ToSymbolKind(string kind) {
			switch (kind) {
				case "Module":
					return LSP.SymbolKind.Module;
				case "Class":
					return LSP.SymbolKind.Class;
				case "Property":
					return LSP.SymbolKind.Property;
				case "Method":
					return LSP.SymbolKind.Method;
				case "Struct":
					return LSP.SymbolKind.Struct;
				case "Variable":
					return LSP.SymbolKind.Variable;
				case "Enum":
					return LSP.SymbolKind.Enum;
				case "EnumMember":
					return LSP.SymbolKind.EnumMember;
				default:
					return LSP.SymbolKind.Object;
			}
		}
	}
}
