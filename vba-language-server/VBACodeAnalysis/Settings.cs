using System;
using System.Collections.Generic;
using System.Text;

namespace VBACodeAnalysis {
	public class Settings {
		//public string NameSpace { get; set; }
		//public Dictionary<string, string> Rewrite { get; set; }
		//public Dictionary<string, string> Restor { get; set; }
		public RewriteSetting RewriteSetting { get; set; }

		public Settings() {
			//NameSpace = "f";
			//Rewrite = new Dictionary<string, string>();
			//Restor = new Dictionary<string, string>();
			//Rewrite.Add("Range", "Ran");
			//Restor.Add("Ran", "Range");
		}

		public void convert() {
			RewriteSetting.convert();

		}
	}

	public class RewriteSetting {
        public string NameSpace { get; set; }
        public List<List<string>> Mapping { get; set; }
        private Dictionary<string, string> Rewrite;
        private Dictionary<string, string> Restore;

		public RewriteSetting() {
			Rewrite = new Dictionary<string, string>();
			Restore = new Dictionary<string, string>();
			Mapping = new List<List<string>>();
		}

		public Dictionary<string, string> getRewriteDict() {
            return Rewrite;
        }
        public Dictionary<string, string> getRestoreDict() {
            return Restore;
        }

        public void convert() {
			Rewrite = new Dictionary<string, string>();
			Restore = new Dictionary<string, string>();
			foreach (var item in Mapping) {
                var src = item[0];
                var dist = item[1];
                Rewrite.Add(src, dist);
                Restore.Add(dist, src);
            }
        }
    }
}
