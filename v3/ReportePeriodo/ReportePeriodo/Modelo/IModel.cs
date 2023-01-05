using ReportePeriodo.Entidad;
using System;
using System.Collections.Generic;
using gth = LibreriaAlimentacion.Entidad;

namespace ReportePeriodo.Modelo
{
    public interface IModel
    {
        string ConexionAccess { get; set; }
        string ConexionSQL { get; set; }
        string ConexionFDB { get; set; }
        string ConexionSIE { get; set; }
        int HoraCorte(ref string mensaje);
        Rancho Rancho(ref string mensaje);
        string Erp_Clave(int ranID, ref string mensaje);
        List<Hoja1> ReporteHoja1(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, out List<CalostroYOrdeña> datosCalostro, ref string mensaje);
        List<gth.IndicadorTeorico> IndicadorProduccion { get; set; }
        List<gth.IndicadorTeorico> IndicadorSecas { get; set; }
        List<gth.IndicadorTeorico> IndicadorReto { get; set; }
        List<gth.IndicadorTeorico> IndicadorCrecimiento { get; set; }
        List<gth.IndicadorTeorico> IndicadorDesarrollo { get; set; }
        List<gth.IndicadorTeorico> IndicadorVaquillas { get; set; }
        void CargarDatosTeoricos(Rancho rancho, DateTime fechaFin, ref string mensaje);
        gth.IndicadorReportePeriodo PromedioProduccion { get; set; }
        gth.IndicadorReportePeriodo PromedioSecas { get; set; }
        gth.IndicadorReportePeriodo PromedioReto { get; set; }
        gth.IndicadorReportePeriodo PromedioCrecimiento { get; set; }
        gth.IndicadorReportePeriodo PromedioDesarrollo { get; set; }
        gth.IndicadorReportePeriodo PromedioVaquillas { get; set; }
        void CargarPromediosDatosAlimentacion(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, ref string mensaje);
        Hoja1 DiferenciaHoja1(Hoja1 promedio, Hoja1 añoAnterior);
        Hoja1 PorcentajeDiferenciaHoja1(Hoja1 diferencia, Hoja1 añoAnterior);
        List<Hoja1> EspaciosEnBlancoHoja1(int renglones);

    }
}
