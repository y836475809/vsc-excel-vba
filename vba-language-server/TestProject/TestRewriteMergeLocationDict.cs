using System.Collections.Generic;
using VBACodeAnalysis;
using Xunit;

namespace TestProject {
	public class TestRewriteMergeLocationDict {
		private Rewrite MakeRewrite() {
			var setting = new RewriteSetting();
			setting.NameSpace = "f";
			return new Rewrite(setting);
		}

		[Fact]
		public void TestMergeDict() {
			var srcDict = new Dictionary<int, List<LocationDiff>> {
				{ 0, new List<LocationDiff>{
					new LocationDiff(0, 0, 1),
					new LocationDiff(0, 5, 1),
					new LocationDiff(0, 20, 1),
				}}
			};
			var dict = new Dictionary<int, List<LocationDiff>> {
				{ 0, new List<LocationDiff>{
					new LocationDiff(0, 8, 1),
					new LocationDiff(0, 10, 1),
				}}
			};
			var rewrite = MakeRewrite();
			rewrite.MergeLocationDiffDict(ref srcDict, dict);

			var predict = new Dictionary<int, List<LocationDiff>> {
				{0, new List<LocationDiff>{
					new LocationDiff(0, 0, 1),
					new LocationDiff(0, 5, 1),
					new LocationDiff(0, 6, 1),
					new LocationDiff(0, 8, 1),
					new LocationDiff(0, 20, 1),
				} },
			};
			Helper.AssertLocationDiffDict(predict, srcDict);
		}

		[Fact]
		public void TestMergeEmptySrcDict() {
			var srcDict = new Dictionary<int, List<LocationDiff>>();
			var dict = new Dictionary<int, List<LocationDiff>> {
				{ 0, new List<LocationDiff>{
					new LocationDiff(0, 8, 1),
					new LocationDiff(0, 10, 1),
				}}
			};
			var rewrite = MakeRewrite();
			rewrite.MergeLocationDiffDict(ref srcDict, dict);

			var predict = new Dictionary<int, List<LocationDiff>> {
				{0, new List<LocationDiff>{
					new LocationDiff(0, 8, 1),
					new LocationDiff(0, 10, 1),
				} },
			};
			Helper.AssertLocationDiffDict(predict, srcDict);
		}

		[Fact]
		public void TestMergeEmptyInDict() {
			var srcDict = new Dictionary<int, List<LocationDiff>> {
				{ 0, new List<LocationDiff>{
					new LocationDiff(0, 8, 1),
					new LocationDiff(0, 10, 1),
				}}
			};
			var dict = new Dictionary<int, List<LocationDiff>>();
			var rewrite = MakeRewrite();
			rewrite.MergeLocationDiffDict(ref srcDict, dict);

			var predict = new Dictionary<int, List<LocationDiff>> {
				{0, new List<LocationDiff>{
					new LocationDiff(0, 8, 1),
					new LocationDiff(0, 10, 1),
				} },
			};
			Helper.AssertLocationDiffDict(predict, srcDict);
		}

		[Fact]
		public void TestMergeDictSort() {
			var srcDict = new Dictionary<int, List<LocationDiff>> {
				{ 0, new List<LocationDiff>{
					new LocationDiff(0, 20, 1),
					new LocationDiff(0, 0, 1),
					new LocationDiff(0, 5, 1),
				}}
			};
			var dict = new Dictionary<int, List<LocationDiff>> {
				{ 0, new List<LocationDiff>{
					new LocationDiff(0, 8, 1),
				}}
			};
			var rewrite = MakeRewrite();
			rewrite.MergeLocationDiffDict(ref srcDict, dict);

			var predict = new Dictionary<int, List<LocationDiff>> {
				{0, new List<LocationDiff>{
					new LocationDiff(0, 0, 1),
					new LocationDiff(0, 5, 1),
					new LocationDiff(0, 6, 1),
					new LocationDiff(0, 20, 1),
				} },
			};
			Helper.AssertLocationDiffDict(predict, srcDict);
		}

		[Fact]
		public void TestMergeDictAgain() {
			var srcDict = new Dictionary<int, List<LocationDiff>> {
				{ 0, new List<LocationDiff>{
					new LocationDiff(0, 0, 1),
					new LocationDiff(0, 5, 1),
					new LocationDiff(0, 20, 1),
				}}
			};
			var dict1 = new Dictionary<int, List<LocationDiff>> {
				{ 0, new List<LocationDiff>{
					new LocationDiff(0, 8, 1),
				}}
			};
			var dict2 = new Dictionary<int, List<LocationDiff>> {
				{ 0, new List<LocationDiff>{
					new LocationDiff(0, 10, 1),
				}}
			};
			var rewrite = MakeRewrite();
			rewrite.MergeLocationDiffDict(ref srcDict, dict1);
			rewrite.MergeLocationDiffDict(ref srcDict, dict2);

			var predict = new Dictionary<int, List<LocationDiff>> {
				{0, new List<LocationDiff>{
					new LocationDiff(0, 0, 1),
					new LocationDiff(0, 5, 1),
					new LocationDiff(0, 6, 1),
					new LocationDiff(0, 7, 1),
					new LocationDiff(0, 20, 1),
				} },
			};
			Helper.AssertLocationDiffDict(predict, srcDict);
		}

		[Fact]
		public void TestMergeDictDup() {
			var srcDict = new Dictionary<int, List<LocationDiff>> {
				{ 0, new List<LocationDiff>{
					new LocationDiff(0, 0, 6),
					new LocationDiff(0, 7, 6),
					new LocationDiff(0, 21, 6),
				}}
			};
			var dict = new Dictionary<int, List<LocationDiff>> {
				{ 0, new List<LocationDiff>{
					new LocationDiff(0, 26, 6),
				}}
			};
			var rewrite = MakeRewrite();
			rewrite.MergeLocationDiffDict(ref srcDict, dict);

			var predict = new Dictionary<int, List<LocationDiff>> {
				{0, new List<LocationDiff>{
					new LocationDiff(0, 0, 6),
					new LocationDiff(0, 7, 6),
					new LocationDiff(0, 14, 6),
					new LocationDiff(0, 21, 6),
				} },
			};
			Helper.AssertLocationDiffDict(predict, srcDict);
		}

		[Fact]
		public void TestMergeDictDupAganin() {
			var srcDict = new Dictionary<int, List<LocationDiff>> {
				{ 0, new List<LocationDiff>{
					new LocationDiff(0, 0, 6),
					new LocationDiff(0, 21, 6),
				}}
			};
			var dict1 = new Dictionary<int, List<LocationDiff>> {
				{ 0, new List<LocationDiff>{
					new LocationDiff(0, 13, 6),
				}}
			};
			var dict2 = new Dictionary<int, List<LocationDiff>> {
				{ 0, new List<LocationDiff>{
					new LocationDiff(0, 26, 6),
				}}
			};
			var rewrite = MakeRewrite();
			rewrite.MergeLocationDiffDict(ref srcDict, dict1);
			rewrite.MergeLocationDiffDict(ref srcDict, dict2);

			var predict = new Dictionary<int, List<LocationDiff>> {
				{0, new List<LocationDiff>{
					new LocationDiff(0, 0, 6),
					new LocationDiff(0, 7, 6),
					new LocationDiff(0, 14, 6),
					new LocationDiff(0, 21, 6),
				} },
			};
			Helper.AssertLocationDiffDict(predict, srcDict);
		}
	}
}
