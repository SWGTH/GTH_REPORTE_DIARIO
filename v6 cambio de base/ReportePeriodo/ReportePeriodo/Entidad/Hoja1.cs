using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportePeriodo.Entidad
{
    public class Hoja1
    {
        public string Dia { get; set; }
        public decimal? Ordeño { get; set; }
        public decimal? Secas { get; set; }
        public decimal? Hato { get; set; }
        public decimal? Porcentaje_Lact { get; set; }
        public decimal? Porcentaje_Prot { get; set; }
        public decimal? Urea { get; set; }
        public decimal? Porcentaje_Grasa { get; set; }
        public decimal? CCS { get; set; }
        public decimal? CTD { get; set; }
        public decimal? DEC { get; set; }

        public decimal? Leche { get; set; }
        public decimal? Antib { get; set; }
        public decimal? Media { get; set; }
        public decimal? Total { get; set; }
        public decimal? DEL { get; set; }
        public decimal? ANT { get; set; }

        public decimal? EA { get; set; }
        public decimal? ILCA { get; set; }
        public decimal? IC { get; set; }
        public decimal? Costo_Litro { get; set; }
        public decimal? MH { get; set; }
        public decimal? Porcentaje_MS { get; set; }
        public decimal? MS { get; set; }
        public decimal? SA { get; set; }
        public decimal? MSS { get; set; }
        public decimal? EAS { get; set; }
        public decimal? Porcentaje_Sob { get; set; }
        public decimal? Costo_Prod { get; set; }
        public decimal? Costo_MS { get; set; }

        public decimal? Cribas_N1 { get; set; }
        public decimal? Cribas_N2 { get; set; }
        public decimal? Cribas_N3 { get; set; }
        public decimal? Cribas_N4 { get; set; }
        
        public decimal? NoID_Ses1 { get; set; }
        public decimal? NoID_Ses2 { get; set; }
        public decimal? NoID_Ses3 { get; set; }
        
        public decimal? SES1 { get; set; }
        public decimal? SES2 { get; set; }
        public decimal? SES3 { get; set; }
        
        public decimal? Porcentaje_Revueltas { get; set; }

        #region valores Colores
        public string Color_ILCA { get; set; }
        public string Color_IC { get; set; }
        public string Color_CostoLitro { get; set; }
        public string Color_MH { get; set; }
        public string Color_PorcentajeMs { get; set; }
        public string Color_MS { get; set; }
        public string Color_CostoProd { get; set; }
        public string Color_CostoMs { get; set; }

        public string Color_Leche { get; set; }
        public string Color_Media { get; set; }
        public string Color_Total { get; set; }

        public string Color_Ses1 { get; set; }
        public string Color_Ses2 { get; set; }
        public string Color_Ses3 { get; set; }

        public string Color_CTD { get; set; }
        public string Color_ANT { get; set; }

        public string Color_HoraSes1 { get; set; }
        public string Color_HoraSes2 { get; set; }
        public string Color_HoraSes3 { get; set; }

        public string Color_Ordeño { get; set; }        

        #endregion
    }
}
