using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace VBALanguageServer {
    public class VbCodeInfo {
        public string VbCode { get; set; }

        public VbCodeInfo() { }

        public VbCodeInfo(string VbCode) {
            this.VbCode = VbCode;
        }
    }

    public class CodeAdapter {
        Dictionary<string, VbCodeInfo> vbCodeInfoDict;

        public CodeAdapter() {
            vbCodeInfoDict = new Dictionary<string, VbCodeInfo>();
        }

        public Dictionary<string, string> getVbCodeDict() {
            var dict = new Dictionary<string, string>();
            foreach(var kv in vbCodeInfoDict) {
                dict[kv.Key] = kv.Value.VbCode;
            }
            return dict;
        }

        public void parse(string filePath, string vbaCode, out VbCodeInfo vbCodeInfo) {
            if(filePath.EndsWith(".d.vb")) {
                vbCodeInfo = new VbCodeInfo {
                    VbCode = vbaCode
                };
                return;
            }

            var headerCount = 0;
            var name = string.Empty;
            StringReader rs = new StringReader(vbaCode);
            while (rs.Peek() >= 0) {
                var text = rs.ReadLine();
                if (name == string.Empty) {
                    var m = Regex.Match(text, @"Attribute VB_Name = ""(.*)""");
                    if (m.Length > 1) {
                        name = m.Groups[1].Value;
                        continue;
                    }
                }
                if (name != string.Empty) {
                    if (!Regex.IsMatch(text, @"^\s*Attribute\s+VB_[A-Za-z]+\s*=.*")) {
                        headerCount++;
                        break;
                    }
                }
                headerCount++;
            }
            var rn = Environment.NewLine;
            var headerLines = vbaCode.Split(rn)[0..headerCount];
            var bodyLines = vbaCode.Split(rn)[headerCount..];
            var body = string.Join(rn, bodyLines);

            var code = string.Empty;
            if (filePath.EndsWith(".cls")) {
                var classLineOffset = 0;
                var lineNum = GetClassAnnotationLineNum(vbaCode);         
                if (lineNum > 0) {
                    bodyLines[lineNum - headerLines.Length] = $"Public Class {name}";
                    body = string.Join(rn, bodyLines);
                    var post = $"{rn}End Class";
                    code = $"{body}{post}";
                } else {
                    classLineOffset = 1;
                    var pre = $"Public Class {name}{rn}";
                    var post = $"{rn}End Class";
                    code = $"{pre}{body}{post}";
                }
                var lineOffset = headerLines.Length - classLineOffset;
                var h = string.Concat(Enumerable.Repeat("\r\n", lineOffset));
                code = $"{h}{code}";
            }
            if (filePath.EndsWith(".bas")) {
                var pre = $"Public Module {name}{rn}";
                var post = $"{rn}End Module";
                code = $"{pre}{body}{post}";
            }
            vbCodeInfo = new VbCodeInfo {
                VbCode = code
            };
        }

        private int GetClassAnnotationLineNum(string vbaCode) {
            var mc = Regex.Match(vbaCode, @"'\s*@class\s+", RegexOptions.IgnoreCase);
            if (mc.Success) {
                var index = mc.Groups[0].Index;
                var code = vbaCode.Substring(0, index);
                var len = code.Length;
                var count = 0;
                for (var i = 0; i < len; i++) {
                    if (code[i] == '\n') {
                        count++;
                    }
                }
                return count;
            }
            return -1;
        }

        public void SetCode(string filePath, string vbaCode) {
            VbCodeInfo vbCodeInfo;
			var cnvEOL = vbaCode.Replace("\r\n", "\n")
				.Replace("\r", "\n").Replace("\n", Environment.NewLine);
			parse(filePath, cnvEOL, out vbCodeInfo);
            vbCodeInfoDict[filePath] = vbCodeInfo;;
        } 

        public bool Has(string filePath) {
            return vbCodeInfoDict.ContainsKey(filePath);
        }

        public VbCodeInfo GetVbCodeInfo(string filePath) {
            return vbCodeInfoDict[filePath];
        }

        public void Delete(string filePath) {
            if (vbCodeInfoDict.ContainsKey(filePath)) {
                vbCodeInfoDict.Remove(filePath);
            }
        }
    }
}
