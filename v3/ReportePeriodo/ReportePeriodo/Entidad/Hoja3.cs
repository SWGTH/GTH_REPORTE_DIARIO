using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportePeriodo.Entidad
{
    public class Hoja3
    {
        public string Dia { get; set; }

        public decimal Jaulas_Vivas { get; set; }
        public decimal Jaulas_Muertas { get; set; }

        public decimal Destete_Vivas { get; set; }
        public decimal Destete_Muertas { get; set; }

        public decimal Vaquillas_Muertas { get; set; }
        public decimal Vaquillas_Urgente { get; set; }
        public decimal Vaquillas_RF { get; set; }
        public decimal Vaquillas_RR { get; set; }
        public decimal Vaquillas_RG { get; set; }
        public decimal Vaquillas_Otros { get; set; }

        public decimal Vacas_Muertas { get; set; }
        public decimal Vacas_Urgente { get; set; }
        public decimal Vacas_RF { get; set; }
        public decimal Vacas_RR { get; set; }
        public decimal Vacas_RG { get; set; }
        public decimal Vacas_Otros { get; set; }

        public decimal Vaquillas_ND { get; set; }
        public decimal Vaquillas_A { get; set; }
        public decimal Vaquillas_Vivos_H { get; set; }
        public decimal Vaquillas_Vivos_M { get; set; }

        public decimal Vacas_ND { get; set; }
        public decimal Vacas_A { get; set; }
        public decimal Vacas_Vivos_H { get; set; }
        public decimal Vacas_Vivos_M { get; set; }

        public decimal SC { get; set; }
        
        public decimal Partos_Vaquillas { get; set; }
        public decimal Partos_Vacas { get; set; }

        public decimal Muertas_Dia { get; set; }
        public decimal Muertas_Noc { get; set; }

        public decimal Diferencia_Calostro { get; set; }
        public decimal Porcentaje_Calostro { get; set; }
        public decimal Calidad_Calostro { get; set; }

    }
}
