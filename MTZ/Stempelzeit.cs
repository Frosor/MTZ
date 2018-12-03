using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MTZ
{
    class Stempelzeit
    {
        public Stempelzeit(string datum, List<string> zeiten)
        {
            this.datum = datum;
            this.zeiten = zeiten;
        }
        public string datum;
        public List<string> zeiten;
    }
}
