using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReportePeriodo.Entity;
using ReportePeriodo.Model;
using ReportePeriodo.View;
using LibreriaGTH.Consumos.Entidad;

namespace ReportePeriodo.Presenter
{
    public class PresentReporter
    {
        ModelReporte _model;
        IView _view;

        public PresentReporter(IView view, string oleDbConnection, string sqlConnection, string fbdConnection)
        {
            this._model = new ModelReporte(oleDbConnection, sqlConnection, fbdConnection);
            this._view = view;
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

        public void CargarDatos(CentroTyp centro,Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            _model.CargarDatos(centro, rancho, fechaInicio, fechaFin);
        }

        public void CargarDatosTeoricos(CentroTyp centro, DateTime fechaInicio, DateTime fechaFin)
        {
            _model.CargarDatosTeoricos(centro, fechaInicio, fechaFin);
        }

        public void CargarDatosAlimentacion(CentroTyp centro, DateTime fechaInicio, DateTime fechaFin)
        {
            _model.CargarDatosAlimentacion(centro, fechaInicio, fechaFin);
        }
    }
}
