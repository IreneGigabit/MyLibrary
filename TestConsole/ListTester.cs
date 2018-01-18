using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace TestConsole {
	class ListTester {
		//private static string BaseDir = System.AppDomain.CurrentDomain.BaseDirectory;//專案路徑
		private static string CurrDir = System.Environment.CurrentDirectory;//Debug路徑

		static void Main(string[] args) {
			DirectoryInfo dir = new DirectoryInfo(CurrDir);
			string BaseDir=dir.Parent.Parent.FullName;

			List<string> myStringLists = new List<string>();
			myStringLists.Add("1");
			myStringLists.Add("2");
			myStringLists.Add("3");
			myStringLists.Add("4");
			myStringLists.Add("5");
			myStringLists.Add("6");
			myStringLists.Add("7");
			myStringLists.Add("8");
			myStringLists.Add("9");
			myStringLists.Add("10");

			//myStringLists.RemoveAt(2-1);
			//myStringLists.Remove("4");

			foreach (string item in myStringLists) {
				Console.WriteLine(item);
			}

			List<string> myStringLists1 = (List<string>)myStringLists.Clone();
			foreach (string item in myStringLists1) {
				Console.WriteLine(item);
			}

			Console.WriteLine(BaseDir);

			Console.ReadKey();
		}
	}

	static class Extensions {
		public static IList<T> Clone<T>(this IList<T> listToClone) where T : ICloneable {
			return listToClone.Select(item => (T)item.Clone()).ToList();
		}
	}
}
