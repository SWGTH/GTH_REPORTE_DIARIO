using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportePeriodo.Entidad
{
    public class Hoja4
    {
       public string Dia { get; set; }

        public decimal? Ubre_MA { get; set; }
        public decimal? Ubre_SL { get; set; }

        public decimal? Metabolicos_FL { get; set; }
        public decimal? Metabolicos_CET { get; set; }

        public decimal? Locomotores_BE { get; set; }
        public decimal? Locomotores_TRA { get; set; }
        public decimal? Locomotores_GA { get; set; }

        public decimal? Digestivos_AC { get; set; }
        public decimal? Digestivos_ES { get; set; }
        public decimal? Digestivos_DI { get; set; }
        public decimal? Digestivos_TI { get; set; }

        public decimal? Reproductivos_RE { get; set; }
        public decimal? Reproductivos_ME { get; set; }
        public decimal? Reproductivos_PIO { get; set; }
        public decimal? Reproductivos_QUI { get; set; }
        public decimal? Reproductivos_CS { get; set; }

        public decimal? Respiratorios_Neu { get; set; }
        
        public decimal? Becerras_Neu { get; set; }
        public decimal? Becerras_Fie { get; set; }
        public decimal? Becerras_Di { get; set; }
        public decimal? Becerras_Conj { get; set; }
        
        public decimal? Vacas_Diag { get; set; }
        public decimal? Vacas_Pren { get; set; }
        public decimal? Vacas_Porcentaje_Pren { get; set; }
        public decimal? Vacas_Vacias { get; set; }
        public decimal? Vacas_Porcentaje_Vacias { get; set; }

        public decimal? Vaquillas_Diag { get; set; }
        public decimal? Vaquillas_Pren { get; set; }
        public decimal? Vaquillas_Porcentaje_Pren { get; set; }
        public decimal? Vaquillas_Vacias { get; set; }
        public decimal? Vaquillas_Porcentaje_Vacias { get; set; }

        public decimal? Abortos_Vaquillas { get; set; }
        public decimal? Abortos_Vacas { get; set; }
    }
}
