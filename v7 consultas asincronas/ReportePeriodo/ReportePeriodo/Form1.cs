using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using LibreriaGTH.Consumos.Entidad;
using ReportePeriodo.Presenter;
using ReportePeriodo.Entity;
using ReportePeriodo.View;
using System.Diagnostics;
using Microsoft.Reporting.WinForms;

namespace ReportePeriodo
{
    public partial class Form1 : Form, IView
    {
        string _fdbConnection = ConfigurationManager.AppSettings["fdbConnection"];
        string _oleDbConnection = ConfigurationManager.AppSettings["oleDbConnection"];
        string _sqlConnection = ConfigurationManager.AppSettings["sqlConnection"];
        bool _proceso;

        Rancho _rancho;
        CentroTyp _centro;
        //PresentReporter _reportPrensenter;
        PresenterReporteDiario _reportPrensenter;

        bool configuraDiario;

        public Form1()
        {
            InitializeComponent();

            Boolean.TryParse(ConfigurationManager.AppSettings["proceso"].ToString(), out _proceso);
            _fdbConnection = _fdbConnection.Replace("$user", "SYSDBA").Replace("$pwd", "masterkey");
            _oleDbConnection = _oleDbConnection.Replace("$pwd", "CiaPrest_");
            _sqlConnection = _sqlConnection.Replace("$usr", "sa").Replace("$pas", "CiaPrest_");

            _rancho = new Rancho();
            _centro = new CentroTyp();
            //_reportPrensenter = new PresentReporter(this, _oleDbConnection, _sqlConnection, _fdbConnection);
            _reportPrensenter = new PresenterReporteDiario(this, _oleDbConnection, _sqlConnection, _fdbConnection);
            configuraDiario = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Process R = Process.GetCurrentProcess();
            R.PriorityClass = ProcessPriorityClass.High;

            _rancho = _reportPrensenter.GetRancho();
            _centro = _reportPrensenter.GetCentro();
            configuraDiario = _reportPrensenter.ConfiguraDiario();

            DateTime fechaMinima = _reportPrensenter.GetFechaMinima();
            DateTime fechaMaxima = _reportPrensenter.GetFechaMaxima();

            monthCalendar1.MinDate = fechaMinima;
            monthCalendar1.MaxDate = fechaMaxima;

            etiEstablo.Text = "ESTABLO: " + _rancho.Nombre.ToUpper();
            if (_proceso)
            {
                DateTime fechaFin = DateTime.Today;
                DateTime fechaInicio = new DateTime(fechaFin.Year, fechaFin.Month, 1);
                Reporte(_centro, _rancho, fechaInicio, fechaFin);
                _reportPrensenter.ValoresConsistentes(_centro, _rancho, fechaFin, configuraDiario);
                Close();
            }

        }

        private void Configurar()
        {

        }     

        private void Reporte(CentroTyp centro, Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            _reportPrensenter.CargarDatos(centro, rancho, fechaInicio, fechaFin);
            _reportPrensenter.GenerarPromedios(rancho, fechaInicio, fechaFin);
            
            _reportPrensenter.LlenarHojas(centro, rancho, fechaInicio, fechaFin);

            List<Hoja1> reporteHoja1 = _reportPrensenter.GetDatosHoja1;
            List<Hoja2> reporteHoja2 = _reportPrensenter.GetDatosHoja2;
            List<Hoja3> reporteHoja3 = _reportPrensenter.GetDatosHoja3;
            List<Hoja4> reporteHoja4 = _reportPrensenter.GetDatosHoja4;
            List<decimal?> listaDecenas = _reportPrensenter.ListaDecenas;
            decimal? novillas = _reportPrensenter.Novillas;
            decimal? vacas = _reportPrensenter.Vacas;

            DataTable datosHoja1 = LibreriaGTH.Utils.ConvertToDataTable<Hoja1>(reporteHoja1);
            DataTable datosHoja2 = LibreriaGTH.Utils.ConvertToDataTable<Hoja2>(reporteHoja2);
            DataTable datosHoja3 = LibreriaGTH.Utils.ConvertToDataTable<Hoja3>(reporteHoja3);
            DataTable datosHoja4 = LibreriaGTH.Utils.ConvertToDataTable<Hoja4>(reporteHoja4);

            decimal decena1 = 0, decena2 = 0, decena3 = 0;

            decimal.TryParse(listaDecenas[0].ToString(), out decena1);
            decimal.TryParse(listaDecenas[1].ToString(), out decena2);
            decimal.TryParse(listaDecenas[2].ToString(), out decena3);

            decimal precioLeche = _reportPrensenter.PrecioLeche;
            decimal precioLecheAñoAnt = _reportPrensenter.PrecioLecheAñoAnt;

            try
            {
                string tituloNoid = rancho.NoIdReal ? "N° ID" : "N° ID + FE";
                string cadenaNovillas = "TOTAL VAQUILLAS: " + (novillas ?? 0).ToString("n0");
                string cadenaVacas = "TOTAL VACAS: " + (vacas ?? 0).ToString("n0");
                string cadenaPrecioLeche = "VENTA (PRECIO DE LA LECHE Actual: @actual / Año Anterior: @anterior)";
                cadenaPrecioLeche = cadenaPrecioLeche.Replace("@actual", precioLeche.ToString("C3")).Replace("@anterior", precioLecheAñoAnt.ToString("C3"));

                ReportParameter[] parameters = new ReportParameter[9];
                parameters[0] = new ReportParameter("EMPRESA", _rancho.Nombre.ToUpper(), true);
                parameters[1] = new ReportParameter("MES", fechaFin.ToString("MMMM yyyy").ToUpper(), true);
                parameters[2] = new ReportParameter("DEC_UNO", decena1.ToString(), true);
                parameters[3] = new ReportParameter("DEC_DOS", decena2.ToString(), true);
                parameters[4] = new ReportParameter("DEC_TRES", decena3.ToString(), true);
                parameters[5] = new ReportParameter("NOVILLAS", cadenaNovillas.ToString(), true);
                parameters[6] = new ReportParameter("VACAS", cadenaVacas.ToString(), true);
                parameters[7] = new ReportParameter("PRECIO_LECHE", cadenaPrecioLeche, true);
                parameters[8] = new ReportParameter("TITULONOID", tituloNoid, true);

                rvDiario.LocalReport.DataSources.Clear();
                ReportDataSource source = new ReportDataSource("DataSet1", datosHoja1);
                rvDiario.LocalReport.DataSources.Add(source);
                source = new ReportDataSource("DataSet2", datosHoja2);
                rvDiario.LocalReport.DataSources.Add(source);
                source = new ReportDataSource("DataSet3", datosHoja3);
                rvDiario.LocalReport.DataSources.Add(source);
                source = new ReportDataSource("DataSet4", datosHoja4);
                rvDiario.LocalReport.DataSources.Add(source);
                rvDiario.LocalReport.SetParameters(parameters);
                rvDiario.LocalReport.Refresh();
                rvDiario.RefreshReport();

                string nombreReporte = @"C:\MOVGANADOAUT\Reportes SIO\Dia.pdf";                
                GTHUtils.SavePDF(rvDiario, nombreReporte);

                if (!_proceso)
                    Process.Start(nombreReporte);

            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            DateTime fechaFin = monthCalendar1.SelectionRange.End.Date;
            DateTime fechaInicio = new DateTime(fechaFin.Year, fechaFin.Month, 1);

            Cursor = Cursors.WaitCursor;
            Reporte(_centro, _rancho, fechaInicio, fechaFin);
            Cursor = Cursors.Default;
            Close();
        }
    }
}
