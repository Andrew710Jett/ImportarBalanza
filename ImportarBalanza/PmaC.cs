using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportarBalanza
{
    class PmaC
    {
        public int IdCia { get; set; }
        public string NombreCia { get; set; }
        public int Periodo { get; set; }
        public string Trimestre { get; set; }
        public int PmaDirecta { get; set; }
        public int PmaCedida { get; set; }
        public int PmaRetenida { get; set; }
        public int ResRiesgosRetenida { get; set; }
        public int PmaDevRetenida { get; set; }
        public int CobExcPerdida { get; set; }
        public int AdqDirecta { get; set; }
        public int SinRetenida { get; set; }
        public int ResTecnico { get; set; }
        public int ResOperacion { get; set; }
        public int GastosOp { get; set; }
        public int ResOperacionAnalog { get; set; }
        public int ProdFinan { get; set; }
        public int OtraReserva { get; set; }
        public int IndCombinado { get; set; }
    }
}
