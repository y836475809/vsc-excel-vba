﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace VBARewrite {
	internal class VBAFunction {
		public string Module { get; set; }
		public List<string> Targets { get; set; }
	}

	internal class VBAPredefined {
		public string Module { get; set; }
		public List<string> Targets { get; set; }
	}

	internal class Settings {
		public VBAFunction VBAFunction { get; set; }
		public VBAPredefined VBAPredefined { get; set; }

		public Settings() {
			VBAFunction = new() {
				Module = "f",
				Targets = [
					"cells",
					"columns",
					"range",
					"workbooks",
					"worksheets",
				]
			};
			VBAPredefined = new VBAPredefined() {
				Module = "ExcelVBAFunctions",
				Targets = [
					"cstr", "cdbl", "cbyte", "cint", "clng", 
					"cbool", "strconv", "val", "createobject"
				]
			};
		}
	}
}
