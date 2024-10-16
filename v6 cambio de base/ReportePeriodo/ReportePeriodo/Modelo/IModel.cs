using ReportePeriodo.Entidad;
using System;
using System.Collections.Generic;
using gthEntity = LibreriaGTH.Consumos.Entidad;

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
        List<Hoja1> ReporteHoja1(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, out List<CalostroYOrdeña> datosCalostro, out List<decimal?> listaDEC, ref string mensaje);
        List<gthEntity.IndicadorReporteTeorico> IndicadorProduccion { get; set; }
        List<gthEntity.IndicadorReporteTeorico> IndicadorSecas { get; set; }
        List<gthEntity.IndicadorReporteTeorico> IndicadorReto { get; set; }
        List<gthEntity.IndicadorReporteTeorico> IndicadorCrecimiento { get; set; }
        List<gthEntity.IndicadorReporteTeorico> IndicadorDesarrollo { get; set; }
        List<gthEntity.IndicadorReporteTeorico> IndicadorVaquillas { get; set; }
        void CargarDatosTeoricos(gthEntity.CentroTyp centro, DateTime fechaInicio, DateTime fechaFin, ref string mensaje);
        gthEntity.IndicadorReportePeriodoTyp PromedioProduccion { get; set; }
        gthEntity.IndicadorReportePeriodoTyp PromedioSecas { get; set; }
        gthEntity.IndicadorReportePeriodoTyp PromedioReto { get; set; }
        gthEntity.IndicadorReportePeriodoTyp PromedioJaulas { get; set; }
        gthEntity.IndicadorReportePeriodoTyp PromedioCrecimiento { get; set; }
        gthEntity.IndicadorReportePeriodoTyp PromedioDesarrollo { get; set; }
        gthEntity.IndicadorReportePeriodoTyp PromedioVaquillas { get; set; }
        gthEntity.IndicadorReportePeriodoTyp PromedioAntProduccion { get; set; }
        gthEntity.IndicadorReportePeriodoTyp PromedioAntSecas { get; set; }
        gthEntity.IndicadorReportePeriodoTyp PromedioAntReto { get; set; }
        gthEntity.IndicadorReportePeriodoTyp PromedioAntJaulas { get; set; }
        gthEntity.IndicadorReportePeriodoTyp PromedioAntCrecimiento { get; set; }
        gthEntity.IndicadorReportePeriodoTyp PromedioAntDesarrollo { get; set; }
        gthEntity.IndicadorReportePeriodoTyp PromedioAntVaquillas { get; set; }
        void CargarPromediosDatosAlimentacion(gthEntity.CentroTyp centro, DateTime fechaInicio, DateTime fechaFin, ref string mensaje);
        Hoja1 PromedioHoja1(Rancho rancho, gthEntity.IndicadorReportePeriodoTyp indicadorOrdeño, DateTime fechaInicio, DateTime fechaFin, ref string mensaje);
        
        Hoja1 DiferenciaHoja1(Hoja1 promedio, Hoja1 añoAnterior);
        Hoja1 PorcentajeDiferenciaHoja1(Hoja1 diferencia, Hoja1 añoAnterior);
        List<Hoja1> EspaciosEnBlancoHoja1(int renglones);
        Hoja1 RenglonTotalHoja1(List<Hoja1> request);
        List<Hoja2> ReporteHoja2(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, ref string mensaje);
        decimal PrecioLeche { get; set; }
        void CargarPrecioLeche(gthEntity.CentroTyp centro, Rancho rancho, DateTime fechaInicio, DateTime fechaFin, ref string mensaje);
        Hoja2 PromedioHoja2(gthEntity.CentroTyp centro, Rancho rancho, gthEntity.IndicadorReportePeriodoTyp indicadorOrdeño, gthEntity.IndicadorReportePeriodoTyp indicadorSecas, gthEntity.IndicadorReportePeriodoTyp indicadorReto,
           gthEntity.IndicadorReportePeriodoTyp indicadorJaulas, gthEntity.IndicadorReportePeriodoTyp indicadorCrecimiento, gthEntity.IndicadorReportePeriodoTyp indicadorDesarrollo,
           gthEntity.IndicadorReportePeriodoTyp indicadorVaquillas, DateTime fechaInicio, DateTime fechaFin, ref string mensaje);
        Hoja2 DiferenciaHoja2(Hoja2 promedio, Hoja2 promedioAñoAnt);
        Hoja2 PorcentajeDiferenciaHoja2(Hoja2 diferencia, Hoja2 promedioAñoAnt);
        List<Hoja2> EspaciosEnBlancoHoja2(int renglones);
        Hoja2 RenglonTotalHoja2(List<Hoja2> request);
        void AsignarColorimetriaHoja2(List<Hoja2> reporte, Utilidad utilidad);
        List<Hoja3> ReporteHoja3(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, ref string mensaje);
        Hoja3 TotalReporteHoja3(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, ref string mensaje);
        Hoja3 DiferenciaReporteHoja3(Hoja3 total, Hoja3 totalAñoAnt, ref string mensaje);
        Hoja3 PorcentajeDiferenciaReporteHoja3(Hoja3 diferencia, Hoja3 totalAñoAnt, ref string mensaje);
        void QuitarCeros(List<Hoja3> reporte);
        void QuitarCeros(Hoja3 item);
        List<Hoja3> EspaciosEnBlancoHoja3(int renglones);
        List<Hoja4> ReporteHoja4(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, ref string mensaje);
        Hoja4 TotalHoja4(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, ref string mensaje);
        Hoja4 DiferenciHoja4(Hoja4 total, Hoja4 totalAñoAnt, ref string mensaje);
        Hoja4 PorcentajeDiferenciHoja4(Hoja4 diferencia, Hoja4 totalAñoAnt, ref string mensaje);
        List<Hoja4> EspaciosEnBlancoHoja4(int renglones);
        void QuitarCeros(List<Hoja4> reporte);
        void QuitarCeros(Hoja4 item);
        void CierreMesCorrecto(int ranId, int horaCorte, string urlWebService, DateTime fechaInicio, DateTime fechaFin, out bool validacionMedicina, out bool validacionAlimento);
        DateTime FechaMaxima(ref string mensaje);
        DateTime FechaMinima(ref string mensaje);
        void ValoresConsistentes(gthEntity.CentroTyp centro, Rancho rancho, DateTime fecha);

    }
}
