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
			var settingRewrite = jsonNode?["rewrite"];
			SettingVBAClassToFunction(settingRewrite?["vba_class_to_function"]);
			SettingVBAPredefined(settingRewrite?["vba_predefined"]);
		}

		private void SettingVBAClassToFunction(System.Text.Json.Nodes.JsonNode jsonNode) {
			var settingVBA = this.RewriteSetting.VBAClassToFunction;
			settingVBA.ModuleName = jsonNode?["module_name"].ToString();
			var vba_classes = jsonNode?["vba_classes"].AsArray();
			foreach (var item in vba_classes) {
				settingVBA.VBAClasses.Add(item.ToString());
			}
		}

		private void SettingVBAPredefined(System.Text.Json.Nodes.JsonNode jsonNode) {
			var setting = this.RewriteSetting.VBAPredefined;
			setting.ModuleName = jsonNode?["module_name"].ToString();
		}
	}

	public class RewriteSetting {
		public VBAClassToFunction VBAClassToFunction { get; set; }
		public VBAPredefined VBAPredefined { get; set; }

		public RewriteSetting() {
			VBAClassToFunction = new VBAClassToFunction();
			VBAPredefined = new VBAPredefined();
		}
	}

	public class VBAClassToFunction {
		public string ModuleName { get; set; }
		public HashSet<string> VBAClasses { get; set; }

		public VBAClassToFunction() {
			VBAClasses = new HashSet<string>();
		}
	}

	public class VBAPredefined {
		public string ModuleName { get; set; }
	}
}
