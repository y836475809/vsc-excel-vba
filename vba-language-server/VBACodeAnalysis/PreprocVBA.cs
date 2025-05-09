﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Antlr4.Runtime;
using Microsoft.CodeAnalysis;
using VBAAntlr;


namespace VBACodeAnalysis {
	using ChangeDict = Dictionary<int, List<ChangeVBA>>;
	using ColumnShiftDict = Dictionary<int, List<ColumnShift>>;
	using LineReMapDict = Dictionary<int, int>;
	using PropertyDict = Dictionary<int, PropertyName>;

	public class ColumnShift(int lineIndex, int startCol, int shiftCol) {
		public int LineIndex = lineIndex;
		public int StartCol = startCol;
		public int ShiftCol = shiftCol;
	}

	public class PropertyName(int lineIndex, string prefix, string text, string asType) {
		public int LineIndex = lineIndex;
		public string Prefix = prefix;
		public string Text = text;
		public string AsType = asType;
	}

	public class ModuleHeader {
		public int LastLineIndex;
		public string VbName;
		private ModuleType _moduleType;

		public ModuleHeader(int lastLineIndex, string vbName, ModuleType type) {
			LastLineIndex = lastLineIndex;
			VbName = vbName.Replace("\"", "");
			_moduleType = type;
		}

		public (string, string) GetStartEnd() {
			if(_moduleType == ModuleType.Cls) {
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
	}

	public class AttributeVBName(int line, int startChara, int endChara, string vbaName) {
		public int Line = line;
		public int StartChara = startChara;
		public int EndChara = endChara;
		public string VBAName = vbaName;
	}

	public class ChangeVBA(int lineIndex, (int, int) repColRange, string text, int startCol, bool enableShift=true) {
		private int _lineIndex = lineIndex;
		private (int, int) _repColRange = repColRange;
		private string _text = text;
		public int StartCol = startCol;
		private bool _enableShift = enableShift;

		public (ColumnShift, string) Apply(string line) {
			var (rStart, rEnd) = _repColRange;
			var text = _text;
			if (rStart < 0 && rEnd < 0) {
				var colShift = new ColumnShift(_lineIndex, StartCol, 0);
				return (colShift, text);
			}
			var t1 = line[0..rStart];
			var t2 = line[rEnd..];
			var t3 = $"{t1}{text}";
			var shiftCol = t3.Length - StartCol;
			
			var rep_line = $"{t3}{t2}";
			if(_enableShift) {
				var colShift = new ColumnShift(_lineIndex, StartCol, shiftCol);
				return (colShift, rep_line);
			} else {
				var colShift = new ColumnShift(_lineIndex, StartCol, 0);
				return (colShift, rep_line);
			}
		}
	}

	public class RewriteVBA : IRewriteVBA {
		private ChangeDict _changeDict;
		private Dictionary<string, PropertyName> _propertyNameDict;
		private List<VBADiagnostic> _ignoreDiagnosticList;

		private string _code;
		private ColumnShiftDict _colShiftDict;
		private LineReMapDict _lineReMapDict;
		private PropertyDict _propertyDict;
		private ModuleHeader _moduleHeader;
		private AttributeVBName _attributeVBName;
		private bool _foundOption;
		private const string OptionExplicitOn = "Option Explicit On";

		public string Code {
			get { return _code; }
		}
		public ColumnShiftDict ColShiftDict {
			get { return _colShiftDict; }
		}
		public LineReMapDict LineReMapDict {
			get { return _lineReMapDict; }
		}

		public PropertyDict PropertyDict {
			get { return _propertyDict; }
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

		public List<VBADiagnostic> IgnoreDiagnosticList {
			get { return _ignoreDiagnosticList; }
		}

		public RewriteVBA() {
			_changeDict = [];
			_propertyNameDict = [];
			_propertyDict = [];
			_ignoreDiagnosticList = [];
			_foundOption = false;
		}

		public void AddChange(int lineIndex, (int, int) repColRange, string text, int startCol, bool enableShift = true) {
			if(!_changeDict.TryGetValue(lineIndex, out List<ChangeVBA> value)) {
				value = ([]);
				_changeDict.Add(lineIndex, value);
			}
			value.Add(new ChangeVBA(lineIndex, repColRange, text, startCol, enableShift));
		}

		public void AddChange(int lineIndex, string text) {
			if (!_changeDict.TryGetValue(lineIndex, out List<ChangeVBA> value)) {
				value = ([]);
				_changeDict.Add(lineIndex, value);
			}
			value.Add(new ChangeVBA(lineIndex, (-1, -1), text, 0));
		}

		public void AddPropertyName(int lineIndex, string prefix, string text, string asType) {
			_propertyDict[lineIndex] = new PropertyName(lineIndex, prefix, text, asType);
			
			if (_propertyNameDict.ContainsKey(text)) {
				var propName = _propertyNameDict[text];
				if(propName.AsType == null && asType != null) {
					_propertyNameDict[text].AsType = asType;
				}
				return;
			}
			_propertyNameDict[text] = 
				new PropertyName(lineIndex, prefix, text, asType);
		}
		public void AddModuleAttribute(int lastLineIndex, string vbName, ModuleType type) {
			_moduleHeader = new ModuleHeader(lastLineIndex, vbName, type);
		}

		public void SetAttributeVBName(int line, int startChara, int endChara, string vbName) {
			_attributeVBName = new AttributeVBName(line, startChara, endChara, vbName);
		}

		public void FoundOption() {
			_foundOption = true;
		}

		public void AddIgnoreDiagnostic((int, int) start, (int, int) end, string text) {
			var (startLinel, startCol) = start;
			var (endLinel, endCol) = end;
			_ignoreDiagnosticList.Add(new() {
				Code = text,
				Start = (startLinel, startCol),
				End = (endLinel, endCol)
			});
		}

		public void ApplyChange(string code) {
			_colShiftDict = [];
			_lineReMapDict = [];

			var lines = code.Split(Environment.NewLine).ToList();

			if (_moduleHeader != null) {
				var headerIndex = 0;
				if (_foundOption) {
					lines[0] = OptionExplicitOn;
					headerIndex = 1;
				}
				var (mStart, mEnd) = _moduleHeader.GetStartEnd();
				lines[headerIndex] = mStart;
				for (int i = headerIndex+1; i <= _moduleHeader.LastLineIndex; i++) {
					lines[i] = "";
				}
				lines.Add(mEnd);

				for (int i = 0; i <= headerIndex; i++) {
					_changeDict.Remove(i);
				}
			}

			ReverseSortChangeVBA();
			foreach (var item in _changeDict) {
				var lineIndex = item.Key;
				var line = lines[lineIndex];
				foreach (var changeVba in item.Value) {
					var (colShift, repLine) = changeVba.Apply(line);
					line = repLine;
					if (colShift.ShiftCol == 0) {
						continue;
					}
					if (!_colShiftDict.TryGetValue(lineIndex, out List<ColumnShift> value)) {
						value = ([]);
						_colShiftDict.Add(lineIndex, value);
					}
					value.Add(colShift);
				}
				lines[lineIndex] = line;
			}
			SortColumnShift(_colShiftDict);

			foreach (var item in _propertyNameDict) {
				var name = item.Key;
				var propName = item.Value;
				var inertIndex = lines.Count - 1;
				if (propName.AsType != null) {
					var asType = propName.AsType;
					var opt = StringComparison.OrdinalIgnoreCase;
					if (string.Equals(propName.AsType, "variant", opt)) {
						asType = "Object";
					}
					var prop = $"Public Property {name} As {asType}";
					lines.Insert(inertIndex, prop);
				} else {
					var prop = $"Public Property {name}";
					lines.Insert(inertIndex, prop);
				}
				_lineReMapDict[inertIndex] = propName.LineIndex;
			}

			_code = string.Join(Environment.NewLine,  lines);
		}

		public void ReverseSortChangeVBA() {
			foreach (var item in _changeDict) {
				item.Value.Sort((a,b) => -(a.StartCol - b.StartCol));
			}
		}
		public void SortColumnShift(ColumnShiftDict dict) {
			foreach (var item in dict) {
				item.Value.Sort((a, b) => a.StartCol - b.StartCol);
			}
		}
	}

	public class PreprocVBA {
		protected Dictionary<string, ColumnShiftDict> _fileColShiftDict;
		protected Dictionary<string, LineReMapDict> _fileLineReMapDict;
		protected Dictionary<string, PropertyDict> _filePropertyDict;
		protected Dictionary<string, List<VBADiagnostic>> _fileDiagnosticDict;
		protected Dictionary<string, List<VBADiagnostic>> _fileIgnoreDiagnosticDict;

		public PreprocVBA() {
			_fileColShiftDict = [];
			_fileLineReMapDict = [];
			_filePropertyDict = [];
			_fileDiagnosticDict = [];
			_fileIgnoreDiagnosticDict = [];
		}

		public int GetColShift(string name, int line, int col) {
			if (!_fileColShiftDict.TryGetValue(name, out ColumnShiftDict dict)) {
				return 0;
			}
			if (!dict.TryGetValue(line, out List<ColumnShift> colShifts)) {
				return 0;
			}
			var colShift = colShifts.Where(x => x.StartCol <= col).Select(x => x.ShiftCol).Sum();
			return colShift;
		}

		public int GetReMapLineIndex(string name, int line) {
			if (!_fileLineReMapDict.TryGetValue(name, out LineReMapDict dict)) {
				return -1;
			}
			if (!dict.TryGetValue(line, out int lineIndex)) {
				return -1;
			}
			return lineIndex;
		}

		public bool TryGetProperty(string name, int line, out string prefix, out string propertyName) {
			prefix = null;
			propertyName = null;
			if (!_filePropertyDict.TryGetValue(name, out PropertyDict dict)) {
				return false;
			}
			if (!dict.TryGetValue(line, out PropertyName prop)) {
				return false;
			}
			prefix = prop.Prefix;
			propertyName = prop.Text;
			return true;
		}

		public List<VBADiagnostic> GetDiagnostics(string name) {
			if (!_fileDiagnosticDict.TryGetValue(name, out List<VBADiagnostic> value)) {
				return [];
			}
			return value;
		}

		public List<VBADiagnostic> GetIgnoreDiagnostics(string name) {
			if (!_fileIgnoreDiagnosticDict.TryGetValue(name, out List<VBADiagnostic> value)) {
				return [];
			}
			return value;
		}

		public string Rewrite(string name, string vbaCode) {
			if (name.EndsWith(".d.vb")) {
				return vbaCode;
			}

			var lexer = new VBALexer(new AntlrInputStream(vbaCode));
			var tokens = new CommonTokenStream(lexer);
			var parser = new VBAParser(tokens);
			lexer.RemoveErrorListeners();
			parser.RemoveErrorListeners();

			var rewriteVBA = new RewriteVBA();
			var nn = new VBAListener(rewriteVBA);
			parser.AddParseListener(nn);
			parser.startRule();

			rewriteVBA.ApplyChange(vbaCode);
			_fileColShiftDict[name] = rewriteVBA.ColShiftDict;
			_fileLineReMapDict[name] = rewriteVBA.LineReMapDict;
			_filePropertyDict[name] = rewriteVBA.PropertyDict;
			SetDiagnosticDict(rewriteVBA, name);
			return rewriteVBA.Code;
		}

		private void SetDiagnosticDict(RewriteVBA rewriteVBA, string fp) {
			var name = Path.GetFileNameWithoutExtension(fp);
			var diagnosticList = rewriteVBA.GetAttributeDiagnosticList(name);
			_fileDiagnosticDict[fp] = [.. diagnosticList];
			_fileIgnoreDiagnosticDict[fp] = rewriteVBA.IgnoreDiagnosticList;
		}
	}
}
