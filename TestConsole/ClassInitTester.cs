using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestConsole
{
	class ClassInitTester
    {
		static void Main(string[] args) {
            Dmt dd = new Dmt("12345");
            Dmt dd1 = new Dmt{ seq = "99999" };

            Console.WriteLine("請按任一鍵離開..");
			Console.ReadKey();
		}
	}

    class Dmt
    {
        public string seq { get; set; }
        public string showseq { get; set; }
        public Dmt() : this("") { }
        public Dmt(string seq) {
            this.seq = seq;
            setseq();
            Console.WriteLine(this.showseq);
        }

        public void setseq() {
            this.showseq = "**" + this.seq + "**";
        }
    }
}
