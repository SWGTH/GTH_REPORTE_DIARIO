using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Windows.Forms;
using System.Configuration;
using ReportePeriodo.Entidad;
using ReportePeriodo.Presentador;
using ReportePeriodo.Vista;
using Microsoft.Reporting.WinForms;
using System.IO;
using System.Diagnostics;

namespace ReportePeriodo
{
    public partial class Form1 : Form, IView
    {

        Rancho _rancho;
        ConfiguracionReporte _configuracionReporte;
        Presenter _presentador;
        string _conexionSIE = ConfigurationManager.AppSettings["conexionSIE"];
        string _conexionFDB = ConfigurationManager.AppSettings["conexionFDB"];
        string _conexionSQL = ConfigurationManager.AppSettings["conexionSQL"];
        string _conexionAccess = ConfigurationManager.AppSettings["conexionAccess"];
        bool _origen;// = Convert.ToBoolean(ConfigurationManager.AppSettings["origen"]);
        string _mensaje;

        public Form1()
        {
            InitializeComponent();
            _rancho = new Rancho();
            _configuracionReporte = new ConfiguracionReporte();
            Boolean.TryParse(ConfigurationManager.AppSettings["origen"].ToString(), out _origen);
            ConfigurarReporte();
            _presentador = new Presenter(this, _configuracionReporte.ConexionAccess, _configuracionReporte.ConexionSQL);
            _mensaje = string.Empty;

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Process R = Process.GetCurrentProcess();
            R.PriorityClass = ProcessPriorityClass.High;

            Configurar();
            if (_origen)
            {
                DateTime fechaFin = DateTime.Today;
                DateTime fechaInicio = new DateTime(fechaFin.Year, fechaFin.Month, 1);
                _presentador.ValoresConsistentes(_rancho, fechaFin);
                Reporte(_rancho, fechaInicio, fechaFin);
                Close();
            }
        }


        private void ConfigurarReporte()
        {
            _configuracionReporte = new ConfiguracionReporte()
            {
                ConexionSQL = _conexionSQL.Replace("$usr", "sa").Replace("$pas", "CiaPrest_"),
                ConexionAccess = _conexionAccess.Replace("$pwd", "CiaPrest_"),
                Origen = _origen
            };
        }

        private void Configurar()
        {
            _rancho = _presentador.Rancho(ref _mensaje);
            _configuracionReporte.ConexionFDB = _conexionFDB.Replace("$tracker", "Tracker" + _rancho.Ran_ID.ToString("0#") + ".FDB").Replace("$user", "SYSDBA").Replace("$pwd", "masterkey");
            _presentador.ConexionFDB = _configuracionReporte.ConexionFDB;
            int timeShiftTracker = _presentador.HoraCorte(ref _mensaje);
            string erpClave = _presentador.Erp_Clave(_rancho.Ran_ID, ref _mensaje);
            _configuracionReporte.ConexionSIE = _conexionSIE.Replace("@", erpClave);
            _presentador.ConexionSIE = _configuracionReporte.ConexionSIE;
            _rancho.Erp = erpClave;
            _rancho.TimeShiftTracker = timeShiftTracker;
            etiEstablo.Text = "ESTABLO: " +  _rancho.Ran_Nombre.ToUpper();
            DateTime fechaMaxima = _presentador.FechaMaxima(ref _mensaje);
            DateTime fechaMinima = _presentador.FechaMinima(ref _mensaje);
            monthCalendar1.MinDate = fechaMinima;
            monthCalendar1.MaxDate = fechaMaxima;
        }

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            DateTime fechaFin = monthCalendar1.SelectionRange.End.Date;
            DateTime fechaInicio = new DateTime(fechaFin.Year, fechaFin.Month, 1);
            Cursor = Cursors.WaitCursor;
            Reporte(_rancho, fechaInicio, fechaFin);
            Cursor = Cursors.Default;
            Close();
        }

        public void Reporte(Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            bool validacionAlimento = true;
            bool validacionMedicina = true;
            List<decimal?> listaDEC = new List<decimal?>();
            string validacionesVisibles = "";

            //DatosLecheProd prueba = _presentador.MejorPromedio(fechaInicio.AddYears(-1), fechaFin.AddYears(-1), ref _mensaje);


            DateTime fechaActAux = DateTime.Today;

            if (fechaFin.Month == fechaActAux.Month && fechaFin.Year == fechaActAux.Year)
            {
                _presentador.CierreMesCorrecto(rancho.Ran_ID, rancho.TimeShiftTracker, _conexionSIE, fechaInicio, fechaFin, out validacionMedicina, out validacionAlimento);
                validacionesVisibles = "si";
            }
            else 
            {
                validacionesVisibles = "no";
            }

            _presentador.CargarDatosTeoricos(rancho, fechaFin, ref _mensaje);
            _presentador.CargarPromediosDatosAlimentacion(rancho, fechaInicio, fechaFin, ref _mensaje);
            _presentador.CargarPrecioLeche(rancho, fechaInicio, fechaFin, ref _mensaje);

            #region hoja1
            List<Hoja1> reporteHoja1 = _presentador.ReporteHoja1(rancho, fechaInicio, fechaFin, out listaDEC, ref _mensaje);
            DataTable dtHoja1 = ListToDataTable(reporteHoja1);
            #endregion

            #region hoja2
            List<Hoja2> reporteHoja2 = _presentador.ReporteHoja2(rancho, fechaInicio, fechaFin, ref _mensaje);
            DataTable dtHoja2 = ListToDataTable(reporteHoja2);
            #endregion

            #region hoja3
            decimal? novillas = 0, vacas = 0;
            List<Hoja3> reporteHoja3 = _presentador.ReporteHoja3(_rancho, fechaInicio, fechaFin, out novillas, out vacas, ref _mensaje);
            DataTable dtHoja3 = ListToDataTable(reporteHoja3);
            #endregion

            #region hoja4
            List<Hoja4> reporteHoja4 = _presentador.ReporteHoja4(_rancho, fechaInicio, fechaFin, ref _mensaje);
            DataTable dtHoja4 = ListToDataTable(reporteHoja4);
            #endregion

            #region Datos Mejor Año
            DatosLecheProd datosMA = _presentador.MejorPromedio(fechaInicio, fechaFin, ref _mensaje);
            #endregion

            decimal dec1 = 0, dec2 = 0, dec3 = 0;

            decimal.TryParse(listaDEC[0].ToString(), out dec1);
            decimal.TryParse(listaDEC[1].ToString(), out dec2);
            decimal.TryParse(listaDEC[2].ToString(), out dec3);

            try 
            {
                decimal valNovillas = novillas == null ? 0 : Convert.ToDecimal(novillas);
                decimal valVacas = vacas == null ? 0 : Convert.ToDecimal(vacas);

                string tituloNoid = _rancho.No_ID_Real ? "N° ID" : "N° ID + FE";
                string cadenaNovillas = "TOTAL VAQUILLAS: " + valNovillas.ToString("n0");
                string cadenaVacas = "TOTAL VACAS: " + valVacas.ToString("n0");

                ReportParameter[] parameters = new ReportParameter[18];
                parameters[0] = new ReportParameter("EMPRESA", _rancho.Ran_Nombre.ToUpper(), true);
                parameters[1] = new ReportParameter("MES", fechaFin.ToString("MMMM yyyy").ToUpper(), true);
                parameters[2] = new ReportParameter("DEC_UNO", dec1.ToString(), true);
                parameters[3] = new ReportParameter("DEC_DOS", dec2.ToString(), true);
                parameters[4] = new ReportParameter("DEC_TRES", dec3.ToString(), true);
                parameters[5] = new ReportParameter("NOVILLAS", cadenaNovillas.ToString(), true);
                parameters[6] = new ReportParameter("VACAS", cadenaVacas.ToString(), true);
                parameters[7] = new ReportParameter("PRECIO_LECHE", "VENTA (PRECIO DE LA LECHE: " + _presentador.PrecioLeche.ToString("C3") + ")", true);
                parameters[8] = new ReportParameter("TITULONOID", tituloNoid, true);
                parameters[9] = new ReportParameter("VALIDACIONALIMENTO", validacionAlimento ? "si" : "no");
                parameters[10] = new ReportParameter("VALIDACIONMEDICINA", validacionMedicina ? "si" : "no");
                parameters[11] = new ReportParameter("VALIDACIONESVISIBLES", validacionesVisibles);

                string parametroLeche = datosMA.Leche != 0 ? datosMA.Leche.ToString("N0"): "";
                string parametroAntib = datosMA.Antib != 0 ? datosMA.Antib.ToString("N0") : ""; 
                string parametroMedia = datosMA.Media != 0 ? datosMA.Media.ToString("N2") : "";
                string parametroTotal = datosMA.Total != 0 ? datosMA.Total.ToString("N0") : "";
                string parametroDel = datosMA.Del != 0 ? datosMA.Del.ToString("N0") : "";
                string parametroAnt = datosMA.Ant != 0 ? datosMA.Ant.ToString("N0") : "";

                parameters[12] = new ReportParameter("VALOR_LECHE", parametroLeche);
                parameters[13] = new ReportParameter("VALOR_ANTIB", parametroAntib);
                parameters[14] = new ReportParameter("VALOR_MEDIA", parametroMedia);
                parameters[15] = new ReportParameter("VALOR_TOTAL", parametroTotal);
                parameters[16] = new ReportParameter("VALOR_DEL", parametroDel);
                parameters[17] = new ReportParameter("VALOR_ANT", parametroAnt);

                rvDiario.LocalReport.DataSources.Clear();
                ReportDataSource source = new ReportDataSource("DataSet1", dtHoja1);
                rvDiario.LocalReport.DataSources.Add(source);
                source = new ReportDataSource("DataSet2", dtHoja2);
                rvDiario.LocalReport.DataSources.Add(source);
                source = new ReportDataSource("DataSet3", dtHoja3);
                rvDiario.LocalReport.DataSources.Add(source);
                source = new ReportDataSource("DataSet4", dtHoja4);
                rvDiario.LocalReport.DataSources.Add(source);
                rvDiario.LocalReport.SetParameters(parameters);
                rvDiario.LocalReport.Refresh();
                rvDiario.RefreshReport();

                string nombreReporte = @"C:\MOVGANADOAUT\Reportes SIO\Dia.pdf";
                GTHUtils.SavePDF(rvDiario, nombreReporte);
                                        
                if (!_origen)
                    Process.Start(nombreReporte);
            }
            catch(Exception ex) { Console.WriteLine(ex.Message); }

        }

        private static DataTable ListToDataTable<T>(IList<T> data)
        {
            DataTable table = new DataTable();

            //special handling for value types and string
            if (typeof(T).IsValueType || typeof(T).Equals(typeof(string)))
            {

                DataColumn dc = new DataColumn("Value", typeof(T));
                table.Columns.Add(dc);
                foreach (T item in data)
                {
                    DataRow dr = table.NewRow();
                    dr[0] = item;
                    table.Rows.Add(dr);
                }
            }
            else
            {
                PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
                foreach (PropertyDescriptor prop in properties)
                {
                    table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
                }
                foreach (T item in data)
                {
                    DataRow row = table.NewRow();
                    foreach (PropertyDescriptor prop in properties)
                    {
                        try
                        {
                            row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                        }
                        catch (Exception ex)
                        {
                            row[prop.Name] = DBNull.Value;
                        }
                    }
                    table.Rows.Add(row);
                }
            }
            return table;
        }

    }
}
