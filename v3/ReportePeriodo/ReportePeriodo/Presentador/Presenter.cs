﻿using ReportePeriodo.Entidad;
using ReportePeriodo.Modelo;
using ReportePeriodo.Vista;
using System;
using System.Collections.Generic;

namespace ReportePeriodo.Presentador
{
    public class Presenter
    {
        Model _model;
        IView _view;

        public Presenter(IView view, string conexionAccess, string conexionSQL)
        {
            this._model = new Model(conexionAccess, conexionSQL);
            this._view = view;
        }

        public string ConexionAccess
        {
            get { return _model.ConexionAccess; }
            set { _model.ConexionAccess = value; }
        }

        public string ConexionSQL
        {
            get { return _model.ConexionSQL; }
            set { _model.ConexionSQL = value; }
        }

        public string ConexionFDB
        {
            get { return _model.ConexionFDB; }
            set { _model.ConexionFDB = value; }
        }

        public string ConexionSIE
        {
            get { return _model.ConexionSIE; }
            set { _model.ConexionSIE = value; }
        }

        public int HoraCorte(ref string mensaje)
        {
            return _model.HoraCorte(ref mensaje);
        }

        public Rancho Rancho(ref string mensaje)
        {
            return _model.Rancho(ref mensaje);
        }

        public string Erp_Clave(int ranID, ref string mensaje)
        {
            return _model.Erp_Clave(ranID, ref mensaje);

        }

        public DateTime FechaMaxima(ref string mensaje)
        {
            return _model.FechaMaxima(ref mensaje);
        }

        public DateTime FechaMinima(ref string mensaje)
        {
            return _model.FechaMinima(ref mensaje);
        }

        public List<Hoja1> ReporteHoja1(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, out List<decimal?> listaDEC, ref string mensaje)
        {
            listaDEC = new List<decimal?>();
            List<CalostroYOrdeña> datosCalostro = new List<CalostroYOrdeña>();
            List<Hoja1> response = _model.ReporteHoja1(rancho, fechaInicio, fechaFin, out datosCalostro, out listaDEC, ref mensaje);
            Hoja1 promedio = _model.PromedioHoja1(rancho, _model.PromedioProduccion, fechaInicio, fechaFin, ref mensaje);
            promedio.Dia = "PROM";
            Hoja1 promedioAñoAnt = _model.PromedioHoja1(rancho, _model.PromedioAntProduccion, fechaInicio.AddYears(-1), fechaFin.AddYears(-1), ref mensaje);
            promedioAñoAnt.Dia = "AÑO ANT";
            Hoja1 diferencia = _model.DiferenciaHoja1(promedio, promedioAñoAnt);
            Hoja1 porcentajeDif = _model.PorcentajeDiferenciaHoja1(diferencia, promedioAñoAnt);
            List<Hoja1> espaciosEnBlanco = _model.EspaciosEnBlancoHoja1(response.Count);

            _model.AsignarColorimetriaHoja1(rancho, response, datosCalostro, promedio, fechaInicio, fechaFin, ref mensaje);
            _model.QuitarCeros(response);
            _model.QuitarCeros(promedio);
            _model.QuitarCeros(promedioAñoAnt);
            _model.QuitarCeros(diferencia);
            _model.QuitarCeros(porcentajeDif);

            response.AddRange(espaciosEnBlanco);
            response.Add(promedio);
            response.Add(promedioAñoAnt);
            response.Add(porcentajeDif);
            response.Add(diferencia);

            return response;
        }

        public void CargarDatosTeoricos(Rancho rancho, DateTime fechaFin, ref string mensaje)
        {
            _model.CargarDatosTeoricos(rancho, fechaFin, ref mensaje);
        }

        public void CargarPromediosDatosAlimentacion(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            _model.CargarPromediosDatosAlimentacion(rancho, fechaInicio, fechaFin, ref mensaje);
        }

        public List<Hoja2> ReporteHoja2(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            mensaje = string.Empty;
            List<Hoja2> response = _model.ReporteHoja2(rancho, fechaInicio, fechaFin, ref mensaje);
            Hoja2 promedio = _model.PromedioHoja2(rancho, _model.PromedioProduccion, _model.PromedioSecas, _model.PromedioReto, _model.PromedioJaulas, _model.PromedioCrecimiento, _model.PromedioDesarrollo, _model.PromedioVaquillas, fechaInicio, fechaFin, ref mensaje);
            promedio.Dia = "PROM";
            Hoja2 promedioAñoAnt = _model.PromedioHoja2(rancho, _model.PromedioAntProduccion, _model.PromedioAntSecas, _model.PromedioAntReto, _model.PromedioAntJaulas, _model.PromedioAntCrecimiento, _model.PromedioAntDesarrollo, _model.PromedioAntVaquillas, fechaInicio.AddYears(-1), fechaFin.AddYears(-1), ref mensaje);
            promedioAñoAnt.Dia = "AÑO ANT";
            Hoja2 diferencia = _model.DiferenciaHoja2(promedio, promedioAñoAnt);
            Hoja2 porcentajeDiferencia = _model.PorcentajeDiferenciaHoja2(diferencia, promedioAñoAnt);
            List<Hoja2> espaciosEnBlanco = _model.EspaciosEnBlancoHoja2(response.Count);
            Utilidad utilidad = new Utilidad();
            utilidad.IXA = promedio.IngresoxAnimal;
            utilidad.CXA = promedio.CostoxAnimal;
            utilidad.PORCENTAJE_C = promedio.Porcentaje_CostoxAnimal;
            utilidad.UXA = promedio.UtilidadxAnimal;
            utilidad.PORCENTAJE_U = promedio.Porcentaje_UtilidadxAnimal;

            _model.AsignarColorimetriaHoja2(response, utilidad);
            _model.QuitarCeros(response);
            _model.QuitarCeros(promedio);
            _model.QuitarCeros(promedioAñoAnt);
            _model.QuitarCeros(diferencia);
            _model.QuitarCeros(porcentajeDiferencia);

            response.AddRange(espaciosEnBlanco);
            response.Add(promedio);
            response.Add(promedioAñoAnt);
            response.Add(porcentajeDiferencia);
            response.Add(diferencia);

            return response;
        }

        public decimal PrecioLeche
        {
            get { return _model.PrecioLeche; }
            set { _model.PrecioLeche = value; }
        }

        public void CargarPrecioLeche(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            _model.CargarPrecioLeche(rancho, fechaInicio, fechaFin, ref mensaje);
        }


        public List<Hoja3> ReporteHoja3(Rancho rancho, DateTime fechaInicio, DateTime fechaFin,
            out decimal? novillas, out decimal? vacas, ref string mensaje)
        {
            mensaje = string.Empty;
            List<Hoja3> response = _model.ReporteHoja3(rancho, fechaInicio, fechaFin, ref mensaje);
            Hoja3 total = _model.TotalReporteHoja3(rancho, fechaInicio, fechaFin, ref mensaje);
            Hoja3 totalAñoAnt = _model.TotalReporteHoja3(rancho, fechaInicio.AddYears(-1), fechaFin.AddYears(-1), ref mensaje);
            totalAñoAnt.Dia = "AÑO ANT";
            Hoja3 diferencia = _model.DiferenciaReporteHoja3(total, totalAñoAnt, ref mensaje);
            Hoja3 porcentajeDiferencia = _model.PorcentajeDiferenciaReporteHoja3(diferencia, totalAñoAnt, ref mensaje);
            List<Hoja3> espaciosEnBlanco = _model.EspaciosEnBlancoHoja3(response.Count);

            novillas = total.Jaulas_Vivas + total.Jaulas_Muertas + total.Destete_Vivas + total.Destete_Muertas + total.Vaquillas_Muertas + total.Vaquillas_Urgente + total.Vaquillas_RF + total.Vaquillas_RR + total.Vaquillas_RG + total.Vaquillas_Otros;

            vacas = total.Vacas_Muertas + total.Vacas_Urgente + total.Vacas_RF + total.Vacas_RR + total.Vacas_RG + total.Vacas_Otros;

            _model.QuitarCeros(response);
            _model.QuitarCeros(total);
            _model.QuitarCeros(totalAñoAnt);
            _model.QuitarCeros(diferencia);
            _model.QuitarCeros(porcentajeDiferencia);

            response.AddRange(espaciosEnBlanco);
            response.Add(total);
            response.Add(totalAñoAnt);
            response.Add(porcentajeDiferencia);
            response.Add(diferencia);

            return response;
        }

        public List<Hoja4> ReporteHoja4(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            List<Hoja4> response = _model.ReporteHoja4(rancho, fechaInicio, fechaFin, ref mensaje);
            Hoja4 total = _model.TotalHoja4(rancho, fechaInicio, fechaFin, ref mensaje);
            Hoja4 totalAñoAnt = _model.TotalHoja4(rancho, fechaInicio.AddYears(-1), fechaFin.AddYears(-1), ref mensaje);
            totalAñoAnt.Dia = "AÑO ANT";
            Hoja4 diferencia = _model.DiferenciHoja4(total, totalAñoAnt, ref mensaje);
            Hoja4 porcentajeDiferencia = _model.PorcentajeDiferenciHoja4(diferencia, totalAñoAnt, ref mensaje);
            List<Hoja4> espaciosEnBlanco = _model.EspaciosEnBlancoHoja4(response.Count);

            _model.QuitarCeros(response);
            _model.QuitarCeros(total);
            _model.QuitarCeros(totalAñoAnt);
            _model.QuitarCeros(diferencia);
            _model.QuitarCeros(porcentajeDiferencia);

            response.AddRange(espaciosEnBlanco);
            response.Add(total);
            response.Add(totalAñoAnt);
            response.Add(porcentajeDiferencia);
            response.Add(diferencia);

            return response;
        }

        public void CierreMesCorrecto(int ranId, int horaCorte, string urlWebService, DateTime fechaInicio, DateTime fechaFin, out bool validacionMedicina, out bool validacionAlimento)
        {
            _model.CierreMesCorrecto(ranId, horaCorte, urlWebService, fechaInicio, fechaFin, out validacionMedicina, out validacionAlimento);
        }

        public void ValoresConsistentes(Rancho rancho, DateTime fecha)
        {
            _model.ValoresConsistentes(rancho, fecha);
        }
    }
}
