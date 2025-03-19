using LibreriaGTH.Consumos.Entidad;
using ReportePeriodo.Entity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportePeriodo.Model.Impl
{
    public interface IModel
    {
        string SqlConnection { get; set; }
        string OleDbConnection { get; set; }
        string FdbConnection { get; set; }
        decimal? PrecioLeche { get; set; }
        decimal? PrecioLecheAñoAnterior { get; set; }

        CentroTyp GetCentro();
        Rancho GetRancho();
        DateTime GetFechaMaxima();
        DateTime GetFechaMinima();
        void CargarDatosTeoricos(CentroTyp centro, DateTime fechaInicio, DateTime fechaFin);
        void CargarPromediosDatosAlimentacion(CentroTyp centro, DateTime fechaInicio, DateTime fechaFin);



    }
}
