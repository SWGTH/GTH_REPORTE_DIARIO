using ReportePeriodo.Entidad;
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

        public List<Hoja1> ReporteHoja1(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            List<CalostroYOrdeña> datosCalostro = new List<CalostroYOrdeña>();
            List<Hoja1> response = _model.ReporteHoja1(rancho, fechaInicio, fechaFin, out datosCalostro, ref mensaje);            
            Hoja1 promedio = _model.PromedioHoja1(rancho, fechaInicio, fechaFin, ref mensaje);
            promedio.Dia = "PROM";
            Hoja1 promedioAñoAnt = _model.PromedioHoja1(rancho, fechaInicio.AddYears(-1), fechaFin.AddYears(-1), ref mensaje);
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

    }
}
