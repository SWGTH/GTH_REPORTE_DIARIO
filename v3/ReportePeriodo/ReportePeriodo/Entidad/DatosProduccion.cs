using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportePeriodo.Entidad
{
    public class DatosProduccion
    {
        public DateTime Fecha { get; set; }
        public decimal? Proteina { get; set; }
        public decimal? Grasa { get; set; }
        public decimal? Urea { get; set; }
        public decimal? CCS { get; set; }
        public decimal? CTD { get; set; }

    }
}
