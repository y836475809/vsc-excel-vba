using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VBAAntlr {
	internal class Util {
		public static bool Contains(string value, List<string> list) {
			return list.Contains(value, StringComparer.OrdinalIgnoreCase);
		}
		public static bool Eq(string value1,  string value2) {
			return string.Equals(value1, value2, StringComparison.OrdinalIgnoreCase);
		}
	}
}
