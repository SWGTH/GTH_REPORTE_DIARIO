using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportePeriodo.Entidad
{
    public class Hoja2
    {
        public string Dia { get; set; }
        public decimal? Jaulas_Inventario { get; set; }
        public decimal? Jaulas_Costo { get; set; }
        public decimal? Crecimiento_Inventario { get; set; }
        public decimal? Crecimiento_MH { get; set; }
        public decimal? Crecimiento_Costo { get; set; }
        public decimal? Crecimiento_PorcentajeMS { get; set; }
        public decimal? Crecimiento_MS { get; set; }
        public decimal? Crecimiento_CostoMS { get; set; }
        public decimal? Desarrollo_Inventario { get; set; }
        public decimal? Desarrollo_MH { get; set; }
        public decimal? Desarrollo_Costo { get; set; }
        public decimal? Desarrollo_PorcentajeMS { get; set; }
        public decimal? Desarrollo_MS { get; set; }
        public decimal? Desarrollo_CostoMS { get; set; }
        public decimal? Vaquillas_Inventario { get; set; }
        public decimal? Vaquillas_MH { get; set; }
        public decimal? Vaquillas_Costo { get; set; }
        public decimal? Vaquillas_PorcentajeMS { get; set; }
        public decimal? Vaquillas_MS { get; set; }
        public decimal? Vaquillas_CostoMS { get; set; }
        public decimal? Secas_Inventario { get; set; }
        public decimal? Secas_MH { get; set; }
        public decimal? Secas_PorcentajeMS { get; set; }
        public decimal? Secas_MS { get; set; }
        public decimal? Secas_SA { get; set; }
        public decimal? Secas_Mss { get; set; }
        public decimal? Secas_PorcentajeSob { get; set; }
        public decimal? Secas_Costo { get; set; }
        public decimal? Secas_CostoMS { get; set; }
        public decimal? Reto_Inventario { get; set; }
        public decimal? Reto_MH { get; set; }
        public decimal? Reto_PorcentajeMS { get; set; }
        public decimal? Reto_MS { get; set; }
        public decimal? Reto_SA { get; set; }
        public decimal? Reto_Mss { get; set; }
        public decimal? Reto_PorcentajeSob { get; set; }
        public decimal? Reto_Costo { get; set; }
        public decimal? Reto_CostoMS { get; set; }
        public decimal? IngresoxAnimal { get; set; }
        public decimal? CostoxAnimal { get; set; }
        public decimal? Porcentaje_CostoxAnimal { get; set; }
        public decimal? Inventario_Total { get; set; }
        public decimal? UtilidadxAnimal { get; set; }
        public decimal? Porcentaje_UtilidadxAnimal { get; set; }

        #region valores colores
        public string Color_MH_Crecimiento { get; set; }
        public string Color_PorcenjeMs_Crecimiento { get; set; }
        public string Color_MS_Crecimiento { get; set; }
        public string Color_CostoMS_Crecimiento { get; set; }
        public string Color_Costo_Crecimiento { get; set; }

        public string Color_MH_Desarrollo { get; set; }
        public string Color_PorcenjeMs_Desarrollo { get; set; }
        public string Color_MS_Desarrollo { get; set; }
        public string Color_CostoMS_Desarrollo { get; set; }
        public string Color_Costo_Desarrollo { get; set; }

        public string Color_MH_Vaquillas { get; set; }
        public string Color_PorcenjeMs_Vaquillas { get; set; }
        public string Color_MS_Vaquillas { get; set; }
        public string Color_CostoMS_Vaquillas { get; set; }
        public string Color_Costo_Vaquillas { get; set; }

        public string Color_MH_Secas { get; set; }
        public string Color_PorcenjeMs_Secas { get; set; }
        public string Color_MS_Secas { get; set; }
        public string Color_CostoMS_Secas { get; set; }
        public string Color_Costo_Secas { get; set; }

        public string Color_MH_Reto { get; set; }
        public string Color_PorcenjeMs_Reto { get; set; }
        public string Color_MS_Reto { get; set; }
        public string Color_CostoMS_Reto { get; set; }
        public string Color_Costo_Reto { get; set; }

        public string Color_IXA { get; set; }
        public string Color_CXA { get; set; }
        public string Color_PorcentajeCXA { get; set; }
        public string Color_UXA { get; set; }
        public string Color_PorcentajeUXA { get; set; }

        #endregion

    }
}
