﻿using ReportePeriodo.Entidad;
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
        List<Hoja1> ReporteHoja1(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, out List<CalostroYOrdeña> datosCalostro, out List<decimal?> listaDEC, ref string mensaje);
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
        gth.IndicadorReportePeriodo PromedioJaulas { get; set; }
        gth.IndicadorReportePeriodo PromedioCrecimiento { get; set; }
        gth.IndicadorReportePeriodo PromedioDesarrollo { get; set; }
        gth.IndicadorReportePeriodo PromedioVaquillas { get; set; }
        gth.IndicadorReportePeriodo PromedioAntProduccion { get; set; }
        gth.IndicadorReportePeriodo PromedioAntSecas { get; set; }
        gth.IndicadorReportePeriodo PromedioAntReto { get; set; }
        gth.IndicadorReportePeriodo PromedioAntJaulas { get; set; }
        gth.IndicadorReportePeriodo PromedioAntCrecimiento { get; set; }
        gth.IndicadorReportePeriodo PromedioAntDesarrollo { get; set; }
        gth.IndicadorReportePeriodo PromedioAntVaquillas { get; set; }
        void CargarPromediosDatosAlimentacion(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, ref string mensaje);
        Hoja1 PromedioHoja1(Rancho rancho, gth.IndicadorReportePeriodo indicadorOrdeño, DateTime fechaInicio, DateTime fechaFin, ref string mensaje);
        
        Hoja1 DiferenciaHoja1(Hoja1 promedio, Hoja1 añoAnterior);
        Hoja1 PorcentajeDiferenciaHoja1(Hoja1 diferencia, Hoja1 añoAnterior);
        List<Hoja1> EspaciosEnBlancoHoja1(int renglones);
        List<Hoja2> ReporteHoja2(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, ref string mensaje);
        decimal PrecioLeche { get; set; }
        void CargarPrecioLeche(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, ref string mensaje);
        Hoja2 PromedioHoja2(Rancho rancho, gth.IndicadorReportePeriodo indicadorOrdeño, gth.IndicadorReportePeriodo indicadorSecas, gth.IndicadorReportePeriodo indicadorReto,
           gth.IndicadorReportePeriodo indicadorJaulas, gth.IndicadorReportePeriodo indicadorCrecimiento, gth.IndicadorReportePeriodo indicadorDesarrollo,
           gth.IndicadorReportePeriodo indicadorVaquillas, DateTime fechaInicio, DateTime fechaFin, ref string mensaje);
        Hoja2 DiferenciaHoja2(Hoja2 promedio, Hoja2 promedioAñoAnt);
        Hoja2 PorcentajeDiferenciaHoja2(Hoja2 diferencia, Hoja2 promedioAñoAnt);
        List<Hoja2> EspaciosEnBlancoHoja2(int renglones);
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
        void ValoresConsistentes(Rancho rancho, DateTime fecha);

    }
}
