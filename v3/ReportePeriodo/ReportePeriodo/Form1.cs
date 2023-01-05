using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using ReportePeriodo.Entidad;
using ReportePeriodo.Presentador;
using ReportePeriodo.Vista;

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
            Configurar();

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
            

        }

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            DateTime fechaFin = monthCalendar1.SelectionRange.End.Date;
            DateTime fechaInicio = new DateTime(fechaFin.Year, fechaFin.Month, 1);

            Reporte(_rancho, fechaInicio, fechaFin);
            
        }



        private void Reporte(Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            _presentador.CargarDatosTeoricos(rancho, fechaFin, ref _mensaje);
            _presentador.CargarPromediosDatosAlimentacion(rancho, fechaInicio, fechaFin, ref _mensaje);


            #region hoja1
            List<Hoja1> reporteHoja1 = _presentador.ReporteHoja1(rancho, fechaInicio, fechaFin, ref _mensaje);
            DataTable dtHoja1 = ListToDataTable(reporteHoja1);
            
            

            
            #endregion

            #region hoja2
            
            #endregion

            #region hoja3
            
            #endregion

            #region hoja4
            
            #endregion
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
