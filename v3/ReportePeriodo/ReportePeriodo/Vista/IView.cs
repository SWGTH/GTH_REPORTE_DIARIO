using ReportePeriodo.Entidad;
using System;

namespace ReportePeriodo.Vista
{
    public interface IView
    {
        void Reporte(Rancho rancho, DateTime fechaInicio, DateTime fechaFin);
    }
}
 