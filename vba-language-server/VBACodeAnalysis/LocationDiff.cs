
namespace VBACodeAnalysis {
	public class LocationDiff {
		public int Line;
		public int Chara;
		public int Diff;

		public LocationDiff(int Line, int Chara, int Diff) {
			this.Line = Line;
			this.Chara = Chara;
			this.Diff = Diff;
		}

		public LocationDiff Clone() {
			return (LocationDiff)MemberwiseClone();
		}
	}
}
