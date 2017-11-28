using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestConsole {
	class DecodeTester {
		static void Main(string[] args) {
			//item.Enable的Type是int,但值不連續
			Console.WriteLine("200".Decode(new Dictionary<string, string> { { "100", "處理中" }, { "200", "成功" }, { "300", "失敗" } }, "無"));
			//Console.WriteLine("200".Decode(new string[] { "100", "處理中", "200", "成功", "300", "失敗", "無" }));
			//"200".Decode(new string[] { "100", "處理中", "200", "成功", "300", "失敗", "無" });
			Console.WriteLine("999".Decode(new string[] { "100", "處理中", "200", "成功", "300", "失敗", "無" }));

			Console.WriteLine("請按任一鍵離開..");
			Console.ReadKey();
		}

	}

	public static class DecodeExt {
		public static string Decode(this string value, IDictionary<string, string> messages, string defaultMessage) {
			if (messages.ContainsKey(value)) {
				return messages[value];
			} else {
				//沒有值就用最後一個Message
				return defaultMessage;
			}
		}

		public static string Decode(this string value, string[] messages) {
			string defMsg = null;
			if (messages.Length > 0) {
				if (messages.Length % 2 == 1) {
					defMsg = messages[messages.Length - 1];
				}

				for (int i = 0; i < messages.Length; i += 2) {
					if (messages[i] == value) {
						return messages[i + 1];
					}
					//Console.WriteLine(messages[i]);
				}

				return defMsg;
			}
			return defMsg;
		}

	}
}
