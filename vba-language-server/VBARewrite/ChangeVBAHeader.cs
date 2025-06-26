using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static VBAAntlr.VBAParser;

namespace VBARewrite {
	public enum ModuleType {
		Cls,
		Bas,
	}

	internal class ChnageVBAHeader {
		public List<ChangeData> ChangeDataList { get; set; }
		public int LastLineIndex;
		public string VbName;
		private ModuleType _moduleType;
		private AttributeVBName _attributeVBName;
		private bool _foundOption;
		private const string OptionExplicitOn = "Option Explicit On";
		
		public ChnageVBAHeader() {
			ChangeDataList = [];
			LastLineIndex = -1;
		}

		public void Change(ref List<string> lines) {
			if (LastLineIndex < 0) {
				return;
			}
			var headerIndex = 0;
			if (_foundOption) {
				lines[0] = OptionExplicitOn;
				headerIndex = 1;
			}
			var (mStart, mEnd) = GetStartEnd();
			lines[headerIndex] = mStart;
			for (int i = headerIndex + 1; i <= LastLineIndex; i++) {
				lines[i] = "";
			}
			lines.Add(mEnd);
		}

		public (string, string) GetStartEnd() {
			if (_moduleType == ModuleType.Cls) {
				return (
					$"Public Class {VbName}",
					$"End Class");
			}
			if (_moduleType == ModuleType.Bas) {
				return (
					$"Public Module {VbName}",
					$"End Module");
			}
			throw new Exception($"{_moduleType}");
		}

		public void GetModuleAttributes(ModuleAttributesContext context) {
			var attrs = context.attributeStmt();
			var s = attrs.Where(x => Util.Eq(x.identifier()[0].GetText(), "VB_Name"));
			if (s.Any()) {
				var lastLineIndex = attrs.Max(x => x.Start.Line) - 1;
				var vbNameIdent = s.First().identifier()[1];
				var vbName = vbNameIdent.GetText();
				var vbNameStart = vbNameIdent.Start;
				var type = attrs.Length > 1 ? ModuleType.Cls : ModuleType.Bas;
				LastLineIndex = lastLineIndex;
				VbName = vbName.Replace("\"", "");
				_moduleType = type;

				if (Util.Eq("VB_Name", VbName)) {
					_attributeVBName = new AttributeVBName(
						vbNameStart.Line,
						vbNameStart.Column, vbNameStart.Column + vbName.Length,
						vbName.Trim('"'));
				}
			}
		}

		public void GetModuleOption(ModuleOptionContext context) {
			var st = context.Start;
			ChangeDataList.Add(new(st.Line - 1, ""));
			_foundOption = true;
		}

		public List<VBADiagnostic> GetAttributeDiagnosticList(string name) {
			if (_attributeVBName == null) {
				return [];
			}
			if (_attributeVBName.VBAName == name) {
				return [];
			}
			var attr = _attributeVBName;
			return [
				new(){
					ID = "CS0103",
					Severity = "Error",
					Message = $"File name is {name}, module name is {attr.VBAName}",
					Start = (attr.Line, attr.StartChara),
					End = (attr.Line, attr.EndChara)
				}
			];
		}
	}
}
