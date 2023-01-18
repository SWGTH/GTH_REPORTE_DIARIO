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
        public double EAProdValor { get; set; }
        public double MsProdValor { get; set; }
        public double MsCrecimientosValor { get; set; }
        public double MsDesarolloValor { get; set; }
        public double MsVaquillasValor { get; set; }
        public double MsSecasValor { get; set; }
        public double MsRetoValor { get; set; }
    }
}
