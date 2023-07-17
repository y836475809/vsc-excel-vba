using System;
using System.Collections.Generic;
using System.Text;

namespace VBACodeAnalysis {
	public class Settings {
		public RewriteSetting RewriteSetting { get; set; }

		public Settings() {
			this.RewriteSetting = new RewriteSetting();
		}

		public void Parse(string jsonStr) {
			var jsonNode = System.Text.Json.Nodes.JsonNode.Parse(jsonStr);
			var rewriteVBA = jsonNode?["rewrite"]?["vba_class_to_function"];

			var settingVBA = this.RewriteSetting.VBAClassToFunction;
			settingVBA.ModuleName = rewriteVBA?["module_name"].ToString();
			var vba_classes = rewriteVBA?["vba_classes"].AsArray();
			foreach (var item in vba_classes) {
				settingVBA.VBAClasses.Add(item.ToString());
			}
		}
	}

	public class RewriteSetting {
		public VBAClassToFunction VBAClassToFunction { get; set; }

		public RewriteSetting() {
			VBAClassToFunction = new VBAClassToFunction();
		}
	}

	public class VBAClassToFunction {
		public string ModuleName { get; set; }
		public HashSet<string> VBAClasses { get; set; }

		public VBAClassToFunction() {
			VBAClasses = new HashSet<string>();
		}
	}
}
