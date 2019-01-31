using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestConsole
{
	class RandomTester
	{

		static void Main(string[] args) {

			Console.WriteLine(GetMix(10));

			Console.WriteLine("請按任一鍵關閉..");
			Console.ReadKey();
		}

		public static string GetMix(int length) {
			Random rnd = new Random();
			string rtnValue = "";
			string str = @"0123456789abcdefghigklmnopqrstuvwxyzABCDEFGHIGKLMNOPQRSTUVWXYZ";

			for (int i = 0; i < length; i++) {
				// 返回數字 
				// rtnValue += rnd.Next(10).ToString(); 

				// 返回小寫字母 
				// rtnValue += str.Substring(10+rnd.Next(26),1); 

				// 返回大寫字母 
				// rtnValue += str.Substring(36+rnd.Next(26),1); 

				// 返回大小寫字母混合 
				// rtnValue += str.Substring(10+rnd.Next(52),1); 

				// 返回小寫字母和數字混合 
				// rtnValue += str.Substring(0 + rnd.Next(36), 1); 

				// 返回大寫字母和數字混合 
				// rtnValue += str.Substring(0 + rnd.Next(36), 1).ToUpper(); 

				// 返回大小寫字母和數字混合 
				rtnValue += str.Substring(0 + rnd.Next(61), 1);
			}
			return rtnValue;
		}
	}
}
