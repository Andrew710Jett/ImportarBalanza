using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportarBalanza
{
    public class Balanza
    {
        public string CodigoCuenta  { get; set;}
        public string Cuenta        { get; set;}
        public string Descripcion   { get; set;}
        public string Ramo          { get; set;}
        //public double BalnceIni     { get; set;}
        //public double TotalDR       { get; set;}
        //public double TotalCR       { get; set;}
        public double EndBalance    { get; set;}
        public string Periodo       { get; set;}
    }                        
}


