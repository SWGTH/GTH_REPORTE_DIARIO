using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportePeriodo.Entidad
{
    public class Cribas
    {
        public DateTime Fecha { get; set; }
        public decimal? Nivel1 { get; set; }
        public decimal? Nivel2 { get; set; }
        public decimal? Nivel3 { get; set; }
        public decimal? Nivel4 { get; set; }
    }
}
