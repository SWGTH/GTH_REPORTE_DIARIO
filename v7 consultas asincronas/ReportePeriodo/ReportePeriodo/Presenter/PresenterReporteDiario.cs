using LibreriaGTH.Consumos.Entidad;
using ReportePeriodo.Entity;
using ReportePeriodo.Model;
using ReportePeriodo.View;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportePeriodo.Presenter
{
    public class PresenterReporteDiario
    {
        ModelReporteDiario _model;
        IView _view;

        List<Hoja1> _datosHoja1;
        List<Hoja2> _datosHoja2;
        List<Hoja3> _datosHoja3;
        List<Hoja4> _datosHoja4;

        decimal? _novillas;
        decimal? _vacas;

        List<decimal?> _listaDecenas;

        public PresenterReporteDiario(IView view, string oleDbConnection, string sqlConnection, string fbdConnection)
        {
            this._model = new ModelReporteDiario(oleDbConnection, sqlConnection, fbdConnection);
            this._view = view;            
        }

        public List<Hoja1> GetDatosHoja1
        {
            get 
            {
                if (_datosHoja1 == null)
                    _datosHoja1 = new List<Hoja1>();

                return _datosHoja1; 
            }
        }

        public List<Hoja2> GetDatosHoja2
        {
            get
            {
                if (_datosHoja2 == null)
                    _datosHoja2 = new List<Hoja2>();

                return _datosHoja2;
            }
        }

        public List<Hoja3> GetDatosHoja3
        {
            get
            {
                if (_datosHoja3 == null)
                    _datosHoja3 = new List<Hoja3>();

                return _datosHoja3;
            }
        }

        public List<Hoja4> GetDatosHoja4
        {
            get
            {
                if (_datosHoja4 == null)
                    _datosHoja4 = new List<Hoja4>();


                return _datosHoja4;
            }
        }

        public List<decimal?> ListaDecenas
        {
            get
            {
                if (_listaDecenas == null)
                    _listaDecenas = new List<decimal?>();

                return _listaDecenas;
            }
        }

        public decimal PrecioLeche
        {
            get
            {
                return _model.PrecioLeche;
            }
        }
        public decimal PrecioLecheAñoAnt
        {
            get
            {
                return _model.PrecioLecheAñoAnt;
            }
        }

        public decimal? Novillas
        {
            get
            {
                if (_novillas == null)
                    _novillas = 0;

                return _novillas;
            }
        }
        public decimal? Vacas
        {
            get
            {
                if (_vacas == null)
                    _vacas = 0;

                return _vacas;
            }
        }

        public CentroTyp GetCentro()
        {
            return _model.GetCentro();
        }

        public Rancho GetRancho()
        {
            return _model.GetRancho();
        }

        public DateTime GetFechaMaxima()
        {
            return _model.GetFechaMaxima();
        }

        public DateTime GetFechaMinima()
        {
            return _model.GetFechaMinima();
        }       

        public void CargarDatos(CentroTyp centro, Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            var hojasTask = new Task(() => { GenerarInformacionHojas(centro, rancho, fechaInicio, fechaFin); });
            var alimentacionTask = new Task(() => { GenerarInformacionAlimentacion(centro, fechaInicio, fechaFin); });            


            hojasTask.Start();
            alimentacionTask.Start();

            var info = new List<Task>
            {
                hojasTask,
                alimentacionTask
            };

            Task.WaitAll(info.ToArray());
        }

        public void LlenarHojas(CentroTyp centro, Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            var hoja1Task = new Task(() => { LlenarHoja1(rancho, fechaInicio, fechaFin); });
            var hoja2Task = new Task(() => { LlenarHoja2(rancho, fechaInicio, fechaFin); });
            var hoja3Task = new Task(() => { LlenarHoja3(rancho, fechaInicio, fechaFin); });
            var hoja4Task = new Task(() => { LlenarHoja4(rancho, fechaInicio, fechaFin); });

            hoja1Task.Start();
            hoja2Task.Start();
            hoja3Task.Start();
            hoja4Task.Start();

            var info = new List<Task>
            {
                hoja1Task,
                hoja2Task,
                hoja3Task,
                hoja4Task
            };

            Task.WaitAll(info.ToArray());

        }

        public void GenerarPromedios(Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            _model.GenerarPromedios(rancho, fechaInicio, fechaFin);
        }


        #region generar informacion
        private async Task<bool> GenerarInformacionHojas(CentroTyp centro, Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            _model.CargarInformacionHojas(centro, rancho, fechaInicio, fechaFin);
            return true;
        }

        private async Task<bool> GenerarInformacionAlimentacion(CentroTyp centro, DateTime fechaInicio, DateTime fechaFin)
        {
            _model.CargarDatosAlimentacion(centro, fechaInicio, fechaFin);
            return true;
        }

        #endregion

        #region generar reportes
        private async Task<bool> LlenarHoja1(Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            List<Hoja1> datos = _model.GetReporteHoja1(rancho, fechaInicio, fechaFin, out _listaDecenas);
            Hoja1 promedio = _model.TotalHoja1;
            Hoja1 promedioAñoAnt = _model.TotalHoja1AñoAnt;
            Hoja1 diferencia = _model.DiferenciaHoja1(promedio, promedioAñoAnt);
            Hoja1 porcentajeDif = _model.PorcentajeDiferenciaHoja1(diferencia, promedioAñoAnt);
            List<Hoja1> espaciosEnBlanco = _model.EspaciosEnBlanco<Hoja1>(datos.Count);
            Hoja1 renglonTotal = _model.RenglonTotalHoja1(datos);
            List<CalostroYOrdeña> datosCalostro = _model.DatosCalostro;
            List<BacteriologiaLeche> datosBacteriologia = _model.DatosBacteriologia;

            _model.AsignarColorimetriaHoja1(rancho, datos, datosCalostro, promedio, fechaInicio, fechaFin, datosBacteriologia);
            _model.QuitarCeros(datos);
            _model.QuitarCeros(promedio);
            _model.QuitarCeros(promedioAñoAnt);
            _model.QuitarCeros(diferencia);
            _model.QuitarCeros(porcentajeDif);
            _model.QuitarCeros(renglonTotal);

            _datosHoja1 = new List<Hoja1>();
            _datosHoja1.AddRange(datos);
            _datosHoja1.AddRange(espaciosEnBlanco);
            _datosHoja1.Add(renglonTotal);
            _datosHoja1.Add(promedio);
            _datosHoja1.Add(promedioAñoAnt);
            _datosHoja1.Add(porcentajeDif);
            _datosHoja1.Add(diferencia);

            return true;
        }

        private async Task<bool> LlenarHoja2(Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            List<Hoja2> datos = _model.GetReporteHoja2(rancho, fechaInicio, fechaFin);
            Hoja2 promedio = _model.TotalHoja2;
            Hoja2 promedioAñoAnt = _model.TotalHoja2AñoAnt;
            Hoja2 diferencia = _model.DiferenciaHoja2(promedio, promedioAñoAnt);
            Hoja2 porcentajeDif = _model.PorcentajeDiferenciaHoja2(diferencia, promedioAñoAnt);
            List<Hoja2> espaciosEnBlanco = _model.EspaciosEnBlanco<Hoja2>(datos.Count);
            Hoja2 renglonTotal = _model.RenglonTotalHoja2(datos);

            Utilidad utilidad = new Utilidad();
            utilidad.IXA = promedio.IngresoxAnimal;
            utilidad.CXA = promedio.CostoxAnimal;
            utilidad.PORCENTAJE_C = promedio.Porcentaje_CostoxAnimal;
            utilidad.UXA = promedio.UtilidadxAnimal;
            utilidad.PORCENTAJE_U = promedio.Porcentaje_UtilidadxAnimal;

            _model.AsignarColorimetriaHoja2(datos, utilidad);
            _model.QuitarCeros(datos);
            _model.QuitarCeros(promedio);
            _model.QuitarCeros(promedioAñoAnt);
            _model.QuitarCeros(diferencia);
            _model.QuitarCeros(porcentajeDif);
            _model.QuitarCeros(renglonTotal);

            _datosHoja2 = new List<Hoja2>();
            _datosHoja2.AddRange(datos);
            _datosHoja2.AddRange(espaciosEnBlanco);
            _datosHoja2.Add(renglonTotal);
            _datosHoja2.Add(promedio);
            _datosHoja2.Add(promedioAñoAnt);
            _datosHoja2.Add(porcentajeDif);
            _datosHoja2.Add(diferencia);

            return true;
        }

        private async Task<bool> LlenarHoja3(Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            List<Hoja3> datos = _model.GetReporteHoja3(rancho, fechaInicio, fechaFin);
            CalostroYOrdeña promedio = _model.PromedioCalostroOrdeño;
            CalostroYOrdeña promedioAnt = _model.PromedioCalostroAñoAnt;

            Hoja3 total = _model.GetTotalReporteHoja3(fechaInicio, fechaFin, promedio, "TOTAL");
            Hoja3 totalAñoAnt = _model.GetTotalReporteHoja3(fechaInicio.AddYears(-1), fechaFin.AddYears(-1), promedioAnt, "AÑO ANT");
            Hoja3 diferencia = _model.DiferenciaHoja3(total, totalAñoAnt);
            Hoja3 porcentajeDif = _model.PorcentajeDiferenciaReporteHoja3(diferencia, totalAñoAnt);
            List<Hoja3> espacionEnBlanco = _model.EspaciosEnBlanco<Hoja3>(datos.Count);

            _novillas = total.Jaulas_Vivas + total.Jaulas_Muertas + total.Destete_Vivas + total.Destete_Muertas + total.Vaquillas_Muertas + total.Vaquillas_Urgente + total.Vaquillas_RF + total.Vaquillas_RR + total.Vaquillas_RG + total.Vaquillas_Otros;

            _vacas = total.Vacas_Muertas + total.Vacas_Urgente + total.Vacas_RF + total.Vacas_RR + total.Vacas_RG + total.Vacas_Otros;

            _model.QuitarCeros(datos);
            _model.QuitarCeros(total);
            _model.QuitarCeros(totalAñoAnt);
            _model.QuitarCeros(diferencia);
            _model.QuitarCeros(porcentajeDif);

            _datosHoja3 = new List<Hoja3>();
            _datosHoja3.AddRange(datos);
            _datosHoja3.AddRange(espacionEnBlanco);
            _datosHoja3.Add(total);
            _datosHoja3.Add(totalAñoAnt);
            _datosHoja3.Add(porcentajeDif);
            _datosHoja3.Add(diferencia);


            return true;
        }

        private async Task<bool> LlenarHoja4(Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            List<Hoja4> datos = _model.GetReporteHoja4(rancho, fechaInicio, fechaFin);
            Hoja4 total = _model.GetTotalHoja4(rancho, fechaInicio, fechaFin, "TOTAL");
            Hoja4 totalAñoAnt = _model.GetTotalHoja4(rancho, fechaInicio.AddYears(-1), fechaFin.AddYears(-1), "AÑO ANT");
            Hoja4 diferencia = _model.DiferenciaHoja4(total, totalAñoAnt);
            Hoja4 porcentajeDif = _model.PorcentajeDiferenciHoja4(diferencia, totalAñoAnt);
            List<Hoja4> espaciosEnBlanco = _model.EspaciosEnBlanco<Hoja4>(datos.Count);

            _model.QuitarCeros(datos);
            _model.QuitarCeros(total);
            _model.QuitarCeros(totalAñoAnt);
            _model.QuitarCeros(diferencia);
            _model.QuitarCeros(porcentajeDif);

            _datosHoja4 = new List<Hoja4>();
            _datosHoja4.AddRange(datos);
            _datosHoja4.AddRange(espaciosEnBlanco);
            _datosHoja4.Add(total);
            _datosHoja4.Add(totalAñoAnt);
            _datosHoja4.Add(porcentajeDif);
            _datosHoja4.Add(diferencia);

            return true;
        }
        #endregion

    }
}
