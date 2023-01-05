using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportePeriodo.Entidad
{
    public class InventarioAfiXDia
    {
        public DateTime Fecha { get; set; }
        public decimal? Ordeño { get; set; }
        public decimal? Secas { get; set; }
        public decimal? Reto { get; set; }
        public decimal? Jaulas { get; set; }
        public decimal? Crecimiento { get; set; }
        public decimal? Desarrollo { get; set; }
        public decimal? Vaquillas { get; set; }
        public decimal? InventarioTotal { get; set; }
    }
}
