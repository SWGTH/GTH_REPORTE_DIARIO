using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportePeriodo.Entidad
{
    public class ValConAlim
    {
        public int Ran_ID { get; set; }
        public bool EAProd { get; set; }
        public bool MsProd { get; set; }
        public bool MsCrecimiento { get; set; }
        public bool MsDesarollo { get; set; }
        public bool MsVaquillas { get; set; }
        public bool MsSecas { get; set; }
        public bool MsReto { get; set; }
        public bool ValConsisitentes { get; set; }
        public decimal? EAProdValor { get; set; }
        public decimal? MsProdValor { get; set; }
        public decimal? MsCrecimientosValor { get; set; }
        public decimal? MsDesarolloValor { get; set; }
        public decimal? MsVaquillasValor { get; set; }
        public decimal? MsSecasValor { get; set; }
        public decimal? MsRetoValor { get; set; }
    }
}
