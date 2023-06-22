using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportePeriodo.Entidad
{
    public class ConfiguracionReporte
    {
        public string ConexionSQL { get; set; }
        public string ConexionAccess { get; set; }
        public string ConexionFDB { get; set; }
        public string ConexionSIE { get; set; }
        public bool Origen { get; set; }
    }
}
