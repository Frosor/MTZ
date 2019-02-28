using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MTZ
{
    class DoppelListe
    {
        public DoppelListe(List<string> zeitenListe, List<string> datumsListe)
        {
            this.zeitenListe = zeitenListe;
            this.datumsListe = datumsListe;
        }
        public DoppelListe()
        {

        }
        public List<string> zeitenListe;
        public List<string> datumsListe;
    }
}
