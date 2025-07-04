﻿using System.Collections.Generic;
using VBACodeAnalysis;
using Xunit;


namespace TestProject {
	public class TestRewriteVBAProperty() {
		[Theory]
		[InlineData("1")]
		[InlineData("2")]
		[InlineData("3")]
		public void TestRewrite(string codeId) {
			var code = Helper.getCode($"test_property{codeId}.bas");
			var rewriter = new TestVBARewriter();
			var actCode = rewriter.Rewrite("test", code);
			var expCode = Helper.getCode($"test_property{codeId}_exp.bas");
			Helper.AssertCode(expCode, actCode);

			ColumnShiftDict expColDict = null;
			LineMapDict expLineMapDict = null;
			if (codeId == "1") {
				expColDict = new ColumnShiftDict {
					{0, new (){ new(0, 20, -4) } },
					{4, new (){ new(4, 20, -5) } },
					{9, new (){ new(9, 20, -5) } },
					{13, new (){ new(13, 20, 5) } },
				};
				expLineMapDict = new LineMapDict {
					{ 16, 9 }
				};
			}
			if (codeId == "2") {
				expColDict = new ColumnShiftDict {
					{0, new (){ new(0, 13, -4) } },
					{4, new (){ new(4, 13, 2) } },
					{9, new (){ new(9, 13, 2) } },
					{13, new (){ new(13, 13, 5) } },
				};
				expLineMapDict = new LineMapDict {
					{ 16, 9 }
				};
			}
			if (codeId == "3") {
				expColDict = new ColumnShiftDict {
					{0, new (){ new(0, 20, -4) } },
					{4, new (){ new(4, 20, -5) } },
					{9, new (){ new(9, 20, -5) } },
					{13, new (){ new(13, 20, -5) } },
				};
				expLineMapDict = new LineMapDict {
					{ 16, 9 },
					{ 17, 13 }
				};
			}

			var actColDict = rewriter.ColDict("test");
			Helper.AssertColumnShiftDict(expColDict, actColDict);
			
			var actLineDict = rewriter.LineMapDict("test");
			Helper.AssertDict(actLineDict, expLineMapDict);
		}

		[Theory]
		[InlineData("1")]
		[InlineData("2")]
		[InlineData("3")]
		public void TestRewriteDynamicArray(string codeId) {
			var filename = "test_property_dynamic_array";
			var code = Helper.getCode($"{filename}{codeId}.bas");
			var rewriter = new TestVBARewriter();
			var actCode = rewriter.Rewrite("test", code);
			var expCode = Helper.getCode($"{filename}{codeId}_exp.bas");
			Helper.AssertCode(expCode, actCode);

			ColumnShiftDict expColDict = null;
			LineMapDict expLineMapDict = null;
			if (codeId == "1") {
				expColDict = new ColumnShiftDict {
					{1, new (){ new(1, 7, 1) } },
					{2, new (){ new(2, 7, 1) } },

					{5, new (){ new(5, 20, 5) } },
					{10, new (){ new(10, 10, 1) } },
					{13, new (){ new(13, 1, 20) } },

					{18, new (){ new(18, 13, 2) } },
					{23, new (){ new(23, 10, 1) } },
					{26, new (){ new(26, 1, 20) } },
					
					{31, new (){ new(31, 13, 2) } },
					{36, new (){ new(36, 10, 1) } },
					{39, new (){ new(39, 1, 20) } },
				};
				expLineMapDict = new LineMapDict {
					{ 43, 18 },
					{ 44, 31 }
				};
			}
			if (codeId == "2") {
				expColDict = new ColumnShiftDict {
					{0, new (){ new(0, 20, 5) } },
					{4, new (){ new(4, 1, 19) } },

					{8, new (){ new(8, 13, 2) } },
					{9, new (){ new(9, 10, 1) } },
					{12, new (){ new(12, 1, 20) } },

					{16, new (){ new(16, 13, 2) } },
					{17, new (){ new(17, 10, 2) } },
					{20, new (){ new(20, 1, 21) } },
				};
				expLineMapDict = new LineMapDict {
					{ 23, 8 },
					{ 24, 16 }
				};
			}
			if (codeId == "3") {
				expColDict = new ColumnShiftDict {
					{1, new (){ new(1, 7, 1) } },

					{5, new (){ new(5, 13, 2) } },

					{12, new (){ new(12, 8, 1) } },
					{13, new (){ new(13, 8, 1) } },

					{15, new (){ new(15, 13, 2) } },
					{16, new (){ new(16, 8, 1) } },
				};
				expLineMapDict = new LineMapDict {
					{ 21, 5 },
					{ 22, 15 }
				};
			}
			var actColDict = rewriter.ColDict("test");
			Helper.AssertColumnShiftDict(expColDict, actColDict);

			var actLineDict = rewriter.LineMapDict("test");
			Helper.AssertDict(actLineDict, expLineMapDict);
		}

		[Theory]
		[InlineData("String", "String", "")]
		[InlineData("String", "String", "()")]
		[InlineData("String", "String", "(2)")]
		[InlineData("String", "String", "(2, 2)")]
		[InlineData("String", "String", "(2, 2, 2)")]

		[InlineData("Variant", "Object ", "")]
		[InlineData("Variant", "Object ", "()")]
		[InlineData("Variant", "Object ", "(2)")]
		[InlineData("Variant", "Object ", "(2, 2)")]
		[InlineData("Variant", "Object ", "(2, 2, 2)")]

		[InlineData("String", "String", "(1 To 2)")]
		[InlineData("String", "String", "(1 To 2, 1 To 2)")]
		[InlineData("String", "String", "(1 To 2, 1 To 2, 1 To 2)")]

		[InlineData("Variant", "Object ", "(1 To 2)")]
		[InlineData("Variant", "Object ", "(1 To 2, 1 To 2)")]
		[InlineData("Variant", "Object ", "(1 To 2, 1 To 2, 1 To 2)")]
		public void TestRewriteAsType(string vbaType, string vbType, string array) {
			var filename = "test_property_astype1";
			var code = string.Format(Helper.getCode($"{filename}.bas"), vbaType, array);
			var rewriter = new TestVBARewriter();
			var actCode = rewriter.Rewrite("test", code);
			var expCode = string.Format(Helper.getCode($"{filename}_exp.bas"), vbType, array);
			Helper.AssertCode(expCode, actCode);
		}
	}
}
