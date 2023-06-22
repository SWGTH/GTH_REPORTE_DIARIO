using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportePeriodo.Entidad
{
    public class ValInventario
    {
        public DateTime Fecha { get; set; }
        public decimal? VacasPreñadas { get; set; }
        public decimal? VacasVacias { get; set; }
        public decimal? VaquillasPreñadas { get; set; }
        public decimal? VaquillasVacias { get; set; }

    }
}
