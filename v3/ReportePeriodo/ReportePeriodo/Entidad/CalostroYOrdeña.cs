using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportePeriodo.Entidad
{
    public class CalostroYOrdeña
    {
        public DateTime Fecha { get; set; }
        public decimal? Porcentaje { get; set; }
        public decimal? Diferencia { get; set; }
        public decimal? Calidad { get; set; }
        public decimal? Horario_Sesion1 { get; set; }
        public decimal? Horario_Sesion2 { get; set; }
        public decimal? Horario_Sesion3 { get; set; }
        public decimal? Capacidad_Ordeño { get; set; }
        public decimal? Porcentaje_Revueltas { get; set; }
        public decimal? Litros_Sesion1 { get; set; }
        public decimal? Litros_Sesion2 { get; set; }
        public decimal? Litros_Sesion3 { get; set; }

    }

}
