using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestConsole
{
	class RoundTester
	{
		static void Main(string[] args) {
			double[] dbl = { 0.4, 0.6, 0.5, 1.5, 0.51, 0.49, -0.51, -0.5, -0.49, 1.25, 1.24, 1.26, 1.35, 1170.975 };
			int len = 2;
			foreach (var item in dbl) {
				Console.WriteLine(string.Format("取{0}位", len));
				//Console.WriteLine(string.Format("Math.Round({0})={1}", item, Math.Round(item, len, MidpointRounding.AwayFromZero)));//xx
				Console.WriteLine(string.Format("Decimal.Round({0})={1}", item, Decimal.Round((decimal)item, len, MidpointRounding.AwayFromZero)));
				//Console.WriteLine(string.Format("string.Format({0})={0:N"+ len + "}", item));//有千分位&補0
				Console.WriteLine(string.Format("string.Format({0})={0:0."+ new string('#', len) + "}", item));
				Console.WriteLine("=======");
			}

			Console.WriteLine("請按任一鍵離開..");
			Console.ReadKey();
		}
	}
}
