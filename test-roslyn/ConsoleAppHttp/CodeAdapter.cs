using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace ConsoleAppServer {
    public class VbCodeInfo {
        public string VbCode { get; set; }
        public int LineOffset { get; set; }
        public int PositionOffset { get; set; }

        public VbCodeInfo() { }

        public VbCodeInfo(string VbCode, int LineOffset, int PositionOffset) {
            this.VbCode = VbCode;
            this.LineOffset = LineOffset;
            this.PositionOffset = PositionOffset;
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
                    VbCode = vbaCode,
                    LineOffset = 0,
                    PositionOffset = 0
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
            var header = string.Join(rn, headerLines);
            var bodyLines = vbaCode.Split(rn)[headerCount..];
            var body = string.Join(rn, bodyLines);

            var code = string.Empty;
            var lineOffset = 0;
            var posOffset = 0;
            if (filePath.EndsWith(".cls")) {
                var pre = $"Public Class {name}{rn}";
                var post = $"{rn}End Class";
                code = $"{pre}{body}{post}";
                lineOffset = headerLines.Length - 1;
                posOffset = header.Length - pre.Length;
            }
            if (filePath.EndsWith(".bas")) {
                var pre = $"Module {name}{rn}";
                var post = $"{rn}End Module";
                code = $"{pre}{body}{post}";
                lineOffset = headerLines.Length - 1;
                posOffset = header.Length - pre.Length;
            }
            vbCodeInfo = new VbCodeInfo {
                VbCode = code,
                LineOffset = lineOffset,
                PositionOffset = posOffset
            };
        }

        public void SetCode(string filePath, string vbaCode) {
            VbCodeInfo vbCodeInfo;
            parse(filePath, vbaCode, out vbCodeInfo);
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
