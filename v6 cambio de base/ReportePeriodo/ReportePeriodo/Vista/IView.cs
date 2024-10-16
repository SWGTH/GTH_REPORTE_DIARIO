using ReportePeriodo.Entidad;
using LibreriaGTH.Consumos.Entidad;
using System;

namespace ReportePeriodo.Vista
{
    public interface IView
    {
        void Reporte(CentroTyp centro, Rancho rancho, DateTime fechaInicio, DateTime fechaFin);
    }
}
 