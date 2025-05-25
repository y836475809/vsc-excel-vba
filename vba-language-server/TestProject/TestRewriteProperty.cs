using System.Collections.Generic;
using Xunit;


namespace TestProject {
	public class TestRewriteProperty() {
		[Theory]
		[InlineData("1")]
		[InlineData("2")]
		public void TestRewrite(string codeId) {
			var code = Helper.getCode($"test_property{codeId}.bas");
			var preprocVBA = new TestPreprocVBA();
			var actCode = preprocVBA.Rewrite("test", code);
			var expCode = Helper.getCode($"test_property{codeId}_exp.bas");
			Helper.AssertCode(expCode, actCode);

			ColumnShiftDict expColDict = null;
			if(codeId == "1") {
				expColDict = new ColumnShiftDict {
					{0, new (){ new(0, 20, -3) } },
					{4, new (){ new(4, 20, -2) } },
					{9, new (){ new(9, 20, -4) } },
					{13, new (){ new(13, 20, 5) } },
				};
			}
			if (codeId == "2") {
				expColDict = new ColumnShiftDict {
					{0, new (){ new(0, 13, -3) } },
					{4, new (){ new(4, 13, 5) } },
					{9, new (){ new(9, 13, 3) } },
					{13, new (){ new(13, 13, 5) } },
				};
			}

			var actColDict = preprocVBA.ColDict["test"];
			Helper.AssertColumnShiftDict(expColDict, actColDict);

			var expLineShiftList = new List<(int, int)>() {     
				(1,  2),
				(12, 1),
				(14, 1),
			};
			var actLineShiftDict = preprocVBA.LineShiftDict["test"];
			Helper.AssertLineShift(expLineShiftList, actLineShiftDict);

			var expLineDict = new LineReMapDict {
				{14 - 2,  9},
			};
			var actLineDict = preprocVBA.LineDict["test"];
			Helper.AssertDict(expLineDict, actLineDict);
		}

		[Theory]
		[InlineData("1")]
		[InlineData("2")]
		public void TestRewriteDynamicArray(string codeId) {
			var filename = "test_property_dynamic_array";
			var code = Helper.getCode($"{filename}{codeId}.bas");
			var preprocVBA = new TestPreprocVBA();
			var actCode = preprocVBA.Rewrite("test", code);
			var expCode = Helper.getCode($"{filename}{codeId}_exp.bas");
			Helper.AssertCode(expCode, actCode);

			ColumnShiftDict expColDict = null;
			LineShiftList expLineShiftList = null;
			LineReMapDict expLineDict = null;
			if (codeId == "1") {
				expColDict = new ColumnShiftDict {
					{1, new (){ new(1, 7, 1) } },
					{2, new (){ new(2, 7, 1) } },

					{5, new (){ new(5, 20, 5) } },
					{10, new (){ new(10, 10, 1) } },
					{13, new (){ new(13, 1, 20) } },
					
					{18, new (){ new(18, 13, 3) } },
					{23, new (){ new(23, 10, 1) } },
					{26, new (){ new(26, 1, 20) } },
					
					{31, new (){ new(31, 13, 3) } },
					{36, new (){ new(36, 10, 1) } },
					{39, new (){ new(39, 1, 20) } },
				};
				expLineShiftList = [
					(6,  1),
					(30, 1),
					(43, 1),
				];
				expLineDict = new LineReMapDict {
					{31 - 1,  18},
					{45 - 2,  31},
				};
			}
			if (codeId == "2") {
				expColDict = new ColumnShiftDict {
					{0, new (){ new(0, 20, 5) } },
					{4, new (){ new(4, 1, 19) } },
					{8, new (){ new(8, 13, 3) } },
					{9, new (){ new(9, 10, 1) } },
					{12, new (){ new(12, 1, 20) } },
					{16, new (){ new(16, 13, 3) } },
					{17, new (){ new(17, 10, 2) } },
					{20, new (){ new(20, 1, 21) } },
				};
				expLineShiftList = [
					(1,  1),
					(15, 1),
					(23, 1),
				];
				expLineDict = new LineReMapDict {
					{16 - 1,  8},
					{25 - 2,  16},
				};
			}
			var actColDict = preprocVBA.ColDict["test"];
			Helper.AssertColumnShiftDict(expColDict, actColDict);

			var actLineShiftDict = preprocVBA.LineShiftDict["test"];
			Helper.AssertLineShift(expLineShiftList, actLineShiftDict);

			var actLineDict = preprocVBA.LineDict["test"];
			Helper.AssertDict(expLineDict, actLineDict);
		}
	}
}
