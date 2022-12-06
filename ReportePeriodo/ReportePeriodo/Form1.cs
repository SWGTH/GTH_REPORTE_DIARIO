using Microsoft.Reporting.WinForms;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Configuration;
using LibreriaAlimentacion;
using bendiciones = LibreriaAlimentacion.Entidad;
using ReportePeriodo.Entidad;
using ReportePeriodo.Controlador;

namespace ReportePeriodo
{
    public partial class Form1 : Form
    {

        OleDbConnection ConnCostos = new OleDbConnection();
        DataTable TablaGlobal = new DataTable();
        Connection conn = new Connection();
        int ran_id, emp_id;
        int ran_sie;
        DateTime fecha_fin;
        DateTime fecha_inicio;
        string ruta;
        bool origen;
        string ranNumero, ranCadena, emp_nombre, emp_codigo;
        bendiciones.IndicadorReportePeriodo prom_ordeño;
        bendiciones.IndicadorReportePeriodo prom_secas;
        bendiciones.IndicadorReportePeriodo prom_reto;
        bendiciones.IndicadorReportePeriodo prom_destete1;
        bendiciones.IndicadorReportePeriodo prom_destete2;
        bendiciones.IndicadorReportePeriodo prom_vquillas;
        List<DatosTeorico> listaProduccion;
        List<DatosTeorico> listaSecas;
        List<DatosTeorico> listaReto;
        List<DatosTeorico> listaDestet1;
        List<DatosTeorico> listaDestete2;
        List<DatosTeorico> listaVaquillas;
        Controlador1 controlador1;
        DateTime fechaFinDTP;
        int timeShifTracker;
        bool pesadores;
        DatosProduccion datosProd;
        Utilidad utilidad;
        List<LecheBacteriologia> listaBacteriologia;
        bool noIdReal;

        public Form1()
        {
            InitializeComponent();
            prom_ordeño = new bendiciones.IndicadorReportePeriodo();
            prom_secas = new bendiciones.IndicadorReportePeriodo();
            prom_reto = new bendiciones.IndicadorReportePeriodo();
            prom_destete1 = new bendiciones.IndicadorReportePeriodo();
            prom_destete2 = new bendiciones.IndicadorReportePeriodo();
            prom_vquillas = new bendiciones.IndicadorReportePeriodo();
            listaProduccion = new List<DatosTeorico>();
            listaSecas = new List<DatosTeorico>();
            listaReto = new List<DatosTeorico>();
            listaDestet1 = new List<DatosTeorico>();
            listaDestete2 = new List<DatosTeorico>();
            listaVaquillas = new List<DatosTeorico>();
            controlador1 = new Controlador1();
            timeShifTracker = 0;
            utilidad = new Utilidad();
            listaBacteriologia = new List<LecheBacteriologia>();
            noIdReal = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            conn.Iniciar();
            ran_id = conn.GetRancho();
            pesadores = conn.GetPesadores();

            //ran_sie = conn.RanchoSie(ran_id); Descomentar en caso de emergencia 
            monthCalendar1.Cursor = Cursors.Hand;
            int _fecha;
            DataTable DtEstablo = new DataTable();
            DtEstablo.Columns.Add("ESTABLO").DataType = System.Type.GetType("System.String");
            string query = "select t2.nombre from RANCHOLOCAL t1, ranchos t2 where t1.RanchoLocal = t2.clave";
            conn.QueryMovsio(query, out DtEstablo);
            label1.Text += DtEstablo.Rows[0][0];
            MaxDate(out _fecha);
            DateTime dtn = new DateTime(1900, 01, 01);
            dtn = dtn.AddDays(_fecha);
            dtn = dtn.AddDays(-2);
            monthCalendar1.MaxDate = dtn;
            MinDate(out _fecha);
            DateTime dtmin = new DateTime(1900, 01, 01);
            dtmin = dtmin.AddDays(_fecha);
            dtmin = dtmin.AddDays(-2);
            monthCalendar1.MinDate = dtmin;

            string mensaje = string.Empty;
            origen = Convert.ToBoolean(ConfigurationManager.AppSettings["origen"]);
            noIdReal = controlador1.NoIdReal(ref mensaje);
            if (origen)
            {
                fecha_inicio = new DateTime(DateTime.Now.Date.Year, DateTime.Now.Date.Month, 1);
                if (File.Exists(@"C:\MOVGANADOAUT\Reportes SIO\DIA.pdf"))
                {
                    File.Delete(@"C:\MOVGANADOAUT\Reportes SIO\DIA.pdf");
                }
                fechaFinDTP = DateTime.Today;
                Reporte(fecha_inicio, DateTime.Now.Date, Horas(), noIdReal);
                Close();
            }

        }
        private int MinDate(out int _fecha)
        {
            _fecha = 0;
            DataTable dt;
            string query = "SELECT MIN(FECHA) FROM VALINVENTARIO";
            conn.QueryMovGanado(query, out dt);
            _fecha = Convert.ToInt32(dt.Rows[0][0]);
            return _fecha;
        }
        private int MaxDate(out int _fecha)
        {
            _fecha = 0;
            DataTable dt;
            string query = "SELECT MAX(FECHA) FROM VALINVENTARIO";
            conn.QueryMovGanado(query, out dt);
            _fecha = Convert.ToInt32(dt.Rows[0][0]);
            return _fecha;
        }

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            monthCalendar1.Enabled = false;
            Cursor = Cursors.WaitCursor;
            fecha_fin = monthCalendar1.SelectionRange.End.Date;
            fecha_inicio = new DateTime(fecha_fin.Year, fecha_fin.Month, 1);
            fechaFinDTP = fecha_fin;
            //bool noidReal = checkNodReal.Checked;
            Reporte(fecha_inicio, fecha_fin, Horas(), noIdReal);
            Cursor = Cursors.Default;
            Close();
        }

        public static int ConvertToJulian(DateTime Date)
        {
            TimeSpan ts = (Date - Convert.ToDateTime("01/01/1900"));
            int julianday = ts.Days + 2;
            return julianday;
        }
        private int Horas()
        {
            int horas = 0;
            /*if (ran_id != 33 && ran_id != 39) //Se agrega el if para covadonga y cañada que no contienen traker 
            {*/
            DataTable dt;
            string query = "select paramvalue from bedrijf_params where name = 'DSTimeShift' ";
            conn.QueryTracker(query, out dt);
            horas = dt.Rows[0][0] != DBNull.Value ? Convert.ToInt32(dt.Rows[0][0]) : 0;
            //}


            return horas;

        }

        public Form1(int ran_id, int emp_id, string emp_nombre)
        {
            InitializeComponent();
            this.ran_id = ran_id;
            this.emp_id = emp_id;
            this.emp_nombre = emp_nombre;
        }

        public Form1(int ran_id, int emp_id, string emp_nombre, bool origen)
        {
            InitializeComponent();
            this.ran_id = ran_id;
            this.emp_id = emp_id;
            this.emp_nombre = emp_nombre;
            this.origen = origen;
        }
        private void GetInfo(bool origen)
        {
            DataTable dt;
            // MODIFICADO
            string query;
            if (origen) // se agrega el || ran == 33 porque covadonga y cañada no tienen tracker 
            {
                query = "SELECT rut_ruta FROM ruta WHERE rut_desc = 'sio'";
                conn.QueryMovGanado(query, out dt);
            }
            else
            {
                query = "select rut_ruta from ruta where ran_id = " + ran_id.ToString();
                conn.QuerySIO(query, out dt);
            }

            ruta = dt.Rows.Count > 0 ? dt.Rows[0][0] != DBNull.Value ? dt.Rows[0][0].ToString() : string.Empty : string.Empty;
        }

        private void Reporte(DateTime inicio, DateTime fin, int horas, bool noIdReal)
        {
            DateTime AñoParaElReporte = inicio;
            GetInfo(origen);
            DateTime fecha_inicioAux = inicio;
            DateTime fecha_finAux = fin;
            fecha_inicio = fecha_inicioAux;
            fecha_fin = fecha_finAux;
            //fecha_fin = monthCalendar1.SelectionRange.End;
            //fecha_inicio = new DateTime(fecha_fin.Year, fecha_fin.Month, 1);

            double dec1 = 0, dec2 = 0, dec3 = 0, _ContadorDec1 = 0, _ContadorDec2 = 0, _ContadorDec3 = 0;
            int dif = 24 + horas, _vacasAux = 0, _vaquillasAux = 0; ;
            int julianaI = ConvertToJulian(fecha_inicioAux);
            int julianaF = ConvertToJulian(fecha_finAux);
            int _FechaFinalEnJulianaAñoAnt = ConvertToJulian(fecha_finAux.AddYears(-1));
            int _FechaInicialEnJulianaAñoAnt = ConvertToJulian(fecha_inicioAux.AddYears(-1));
            string _Novillas = "TOTAL VAQUILLAS: ", _Vacas = "TOTAL VACAS: ";
            inicio = horas != 0 ? inicio.AddHours(horas) : inicio;
            fin = horas != 0 ? fin = fin.AddHours(dif) : fin;
            DataTable dtH1, dtH2, dtH2AUX, dtH1AUX;
            bool validacionAlimento = true;
            bool validacionMedicina = true;
            DateTime hoyValida = DateTime.Today;

            if (hoyValida.Year == fecha_finAux.Year && hoyValida.Month == fecha_finAux.Month)
                controlador1.CierreMesCorrecto(ran_id, horas, fecha_inicioAux, fecha_finAux, out validacionMedicina, out validacionAlimento);


            Hoja1(julianaI, julianaF, noIdReal, out dtH1);
            //Se saca una segunda tabla de la primera hoja para poder sacar los ultimos renglones 
            Hoja1(_FechaInicialEnJulianaAñoAnt, _FechaFinalEnJulianaAñoAnt, noIdReal, out dtH1AUX);

            //if (ran_id != 33 && ran_id != 39)// validamos que no sea ni covadonga ni la cañada ya que busca valores de SQL
            //{
            if (origen)
            {
                ValoresConsisitentesAlimentacion(fin);
            }
            if (fin >= new DateTime(2021, 01, 01))
            {
                ValoresReporteDiarioProduccion(fecha_inicioAux, fecha_finAux, dtH1);
                //ValoresReporteDiarioProduccion(inicio.AddYears(-1), fin.AddYears(-1), dtH1AUX); /*(Descomentar en tres meses 17 / 12 / 2021)*/
            }


            //}

            for (int i = 0; i < 31; i++)
            {
                if (dtH1.Rows[i][8] == DBNull.Value) { }
                else
                {
                    if (i <= 9)
                    {
                        dec1 = dec1 + Convert.ToDouble(dtH1.Rows[i][8]);
                        _ContadorDec1++;
                    }
                    if (i > 9 && i <= 19)
                    {
                        dec2 = dec2 + Convert.ToDouble(dtH1.Rows[i][8]);
                        _ContadorDec2++;
                    }
                    if (i > 19)
                    {
                        dec3 = dec3 + Convert.ToDouble(dtH1.Rows[i][8]);
                        _ContadorDec3++;
                    }
                }
            }
            if (dec1 != 0) { dec1 = dec1 / _ContadorDec1; }
            if (dec2 != 0) { dec2 = dec2 / _ContadorDec2; }
            if (dec3 != 0) { dec3 = dec3 / _ContadorDec3; }
            dtH1.Rows[32][41] = dtH1.Rows[32][8];
            for (int i = 1; i < dtH1AUX.Columns.Count; i++)
            {
                if (dtH1AUX.Rows[32][i] == DBNull.Value)//Validamos que no haya un valor nulo, en caso de ser nulo no hace nada 
                {

                }
                else//De otro modo hace la operación para sacar el promedio
                {
                    dtH1.Rows[33][i] = dtH1AUX.Rows[32][i];

                }
            }
            dtH1.Rows[33][41] = dtH1AUX.Rows[32][8];
            //Una vez que tenemos los promedios del año actual y del año anterior sacamos las diferencias
            for (int i = 1; i <= 41; i++)
            {
                //Primero pregunta si el valor es cero o bien nulo, en caso de ser así significa que no hay nada contra lo cual comparar por lo que la diferencia será el
                //mismo valor que el del año pasado.
                if (dtH1.Rows[32][i] == DBNull.Value || Convert.ToString(dtH1.Rows[32][i]) == "NaN")
                {
                    if (dtH1.Rows[33][i] != DBNull.Value)
                    {
                        dtH1.Rows[35][i] = Convert.ToDouble(dtH1.Rows[33][i]) * -1;
                        dtH1.Rows[34][i] = DBNull.Value;
                    }
                }
                else
                {
                    if (dtH1.Rows[33][i] == DBNull.Value || Convert.ToDouble(dtH1.Rows[32][i]) == 0 || Convert.ToString(dtH1.Rows[33][i]) == "NaN")//En caso de tener valor preguntamos si el del año pasado tiene valor, si es nulo no hace nada 
                    {

                    }
                    else//Pero si tiene valor hacemos ambas diferencias.
                    {
                        dtH1.Rows[35][i] = (Convert.ToDouble(dtH1.Rows[32][i]) - Convert.ToDouble(dtH1.Rows[33][i])) * 1;
                        dtH1.Rows[34][i] = Convert.ToDouble(dtH1.Rows[33][i]) != 0 ? (Convert.ToDouble(dtH1.Rows[35][i]) / Convert.ToDouble(dtH1.Rows[33][i])) * 100 : 0;
                        if (Convert.ToDouble(dtH1.Rows[35][i]) > 0)
                        {
                            dtH1.Rows[34][i] = Math.Abs(Convert.ToDouble(dtH1.Rows[34][i]));
                        }
                    }
                }
            }
            //En caso del valo que regrese sea NaN, ponemos un valor nulo en el data table 
            for (int i = 31; i < 36; i++)
            {
                for (int j = 25; j < 40; j++)
                {
                    if (Convert.ToString(dtH1.Rows[32][j]) == "NaN" || Convert.ToString(dtH1.Rows[33][j]) == "NaN" || Convert.ToString(dtH1.Rows[35][j]) == "NaN" || Convert.ToString(dtH1.Rows[32][j]) == "NeuN" || Convert.ToString(dtH1.Rows[33][j]) == "NeuN" || Convert.ToString(dtH1.Rows[35][j]) == "NeuN")
                    {
                        dtH1.Rows[i][j] = DBNull.Value;
                    }
                }
            }
            //Se crea una tabla copia global de la hoja uno con los valores del año actual que usaremos en la hoja 2
            ColumnasDTH1(out TablaGlobal);
            for (int i = 0; i < dtH1.Rows.Count; i++)
            {
                TablaGlobal.Rows.Add();
                for (int j = 0; j < dtH1.Columns.Count; j++)
                {
                    TablaGlobal.Rows[i][j] = dtH1.Rows[i][j];
                }
            }

            //Sacamos las diferencias de la hoja dos, aquí no hay que hacer nada de lo anterior ya que el método que trae los valores del año anterior sí funciona
            //por lo que solo hay que ir por los valores y hacer las operaciones.
            Hoja2(julianaI, julianaF, out dtH2);
            Hoja2(_FechaInicialEnJulianaAñoAnt, _FechaFinalEnJulianaAñoAnt, out dtH2AUX);
            //if (ran_id != 33 && ran_id != 39)// validamos que no sea ni covadonga ni la cañada ya que busca valores de SQL
            //{
            if (fin >= new DateTime(2021, 01, 01))
            {
                ValoresReporteDiarioHojaDos(fecha_inicioAux, fecha_finAux, dtH2);
                //ValoresReporteDiarioHojaDos(inicio.AddYears(-1), fin.AddYears(-1), dtH2AUX); /*(Descomentar en tres meses 17 / 12 / 2021)*/
            }
            //}
            for (int i = 1; i < 50; i++)
            {
                dtH2.Rows[33][i] = dtH2AUX.Rows[32][i];
                if (Convert.ToString(dtH2.Rows[32][i]) == "NaN" || Convert.ToString(dtH2.Rows[32][i]) == "NeuN")
                {
                    dtH2.Rows[32][i] = 0;
                }
                if (dtH2.Rows[32][i] == DBNull.Value || Convert.ToDouble(dtH2.Rows[32][i]) == 0)
                {
                    if (dtH2.Rows[33][i] != DBNull.Value)
                    {
                        dtH2.Rows[35][i] = Convert.ToDouble(dtH2.Rows[33][i]) * -1;
                        dtH2.Rows[34][i] = DBNull.Value;
                    }
                }
                else
                {
                    if (dtH2.Rows[33][i] == DBNull.Value)
                    {

                    }
                    else
                    {
                        dtH2.Rows[35][i] = (Convert.ToDouble(dtH2.Rows[33][i]) - Convert.ToDouble(dtH2.Rows[32][i])) * -1;
                        dtH2.Rows[34][i] = Convert.ToDouble(dtH2.Rows[33][i]) != 0 ? (Convert.ToDouble(dtH2.Rows[35][i]) / Convert.ToDouble(dtH2.Rows[33][i])) * 100 : 0;
                        if (Convert.ToDouble(dtH2.Rows[35][i]) > 0)
                        {
                            dtH2.Rows[34][i] = Math.Abs(Convert.ToDouble(dtH2.Rows[34][i]));
                        }
                    }
                }
                dtH2.Rows[33][i] = Convert.ToString(dtH2.Rows[33][i]) == "NaN" ? DBNull.Value : dtH2.Rows[33][i];
                dtH2.Rows[34][i] = Convert.ToString(dtH2.Rows[34][i]) == "NaN" ? DBNull.Value : dtH2.Rows[34][i];
                dtH2.Rows[35][i] = Convert.ToString(dtH2.Rows[35][i]) == "NaN" ? DBNull.Value : dtH2.Rows[35][i];
            }

            DataTable _DtAux = new DataTable(), _DtSR = new DataTable();
            InfoDPA(out _DtAux);
            //Con la información de InfoDPA sacamos el total de vacas y de vaquilas que irá en los parametros de la hoja dos.
            for (int i = 1; i < 11; i++)
            {
                if (_DtAux.Rows[31][i] == DBNull.Value) {/*validamos valores nulos}*/}
                else
                {
                    _vaquillasAux = _vaquillasAux + Convert.ToInt32(_DtAux.Rows[31][i]);
                }
            }
            for (int i = 11; i < 17; i++)
            {
                if (_DtAux.Rows[31][i] == DBNull.Value) {/*validamos valores nulos*/}
                else
                {
                    _vacasAux = _vacasAux + Convert.ToInt32(_DtAux.Rows[31][i]);
                }
            }
            _Novillas = _Novillas + Convert.ToString(_vaquillasAux);
            _Vacas = _Vacas + Convert.ToString(_vacasAux);
            InfoSR(out _DtSR);


            DataTable DtEstablo = new DataTable();
            DtEstablo.Columns.Add("ESTABLO").DataType = System.Type.GetType("System.String");
            string query = "select t2.nombre from RANCHOLOCAL t1, ranchos t2 where t1.RanchoLocal = t2.clave";
            conn.QueryMovsio(query, out DtEstablo);
            double precioBase = 8.895;

            //Ajuste para saber el precio de la leche dependiendo si es del mes actual o meses pasados
            DateTime FechaReferencia = DateTime.Today;
            string validacionesVisibles = "";

            if (FechaReferencia.Month == inicio.AddDays(1).Month && FechaReferencia.Year == inicio.Year)
            {
                DataTable dtPrecioLeche = new DataTable();
                query = "SELECT* From PRECIOLECHE";
                conn.QueryMovsio(query, out dtPrecioLeche);
                int tipoLeche = Convert.ToInt32(dtPrecioLeche.Rows[0]["TIPOCALCULO"]);
                if (tipoLeche == 2)
                {
                    precioBase = Convert.ToDouble(dtPrecioLeche.Rows[0]["CALCULOUSUARIO"]);
                }
                if (tipoLeche == 1)
                {
                    precioBase = Convert.ToDouble(dtPrecioLeche.Rows[0]["CALCULOLALA"]);
                }
                if (tipoLeche == 3)
                {
                    precioBase = Convert.ToDouble(dtPrecioLeche.Rows[0]["PRECIOUSUARIO"]);
                }
                validacionesVisibles = "si";
            }
            else
            {
                List<bendiciones.ReportePeriodo> listaIngredientes = new List<bendiciones.ReportePeriodo>();
                List<bendiciones.ReportePeriodo> listaForraje = new List<bendiciones.ReportePeriodo>();
                List<bendiciones.ReportePeriodo> listaAlimento = new List<bendiciones.ReportePeriodo>();
                bendiciones.ReportePeriodo sobrante = new bendiciones.ReportePeriodo();
                bendiciones.IndicadorReportePeriodo indicadoresProduccion = new bendiciones.IndicadorReportePeriodo();
                GTH.ReportePeriodo(ran_id.ToString(), horas, "10,11,12,13", inicio.Date, fin.Date, out listaIngredientes, out indicadoresProduccion, out sobrante);
                precioBase = (indicadoresProduccion.IC_PRODUCCION + indicadoresProduccion.COSTO) / indicadoresProduccion.MEDIA;
                validacionesVisibles = "no";
            }

            Console.WriteLine("Ordeño: " + prom_ordeño.MS.ToString());
            Console.WriteLine("Secas: " + prom_secas.MS.ToString());
            Console.WriteLine("Reto: " + prom_reto.MS.ToString());
            Console.WriteLine("Destete1: " + prom_destete1.MS.ToString());
            Console.WriteLine("Destete2: " + prom_destete2.MS.ToString());
            Console.WriteLine("Vaquillas: " + prom_vquillas.MS.ToString());

            string[] meses = new string[] { "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE" };
            string tituloNoid = noIdReal ? "N° ID" : "N° ID + FE";
            ReportParameter[] parameters = new ReportParameter[12];
            parameters[0] = new ReportParameter("EMPRESA", Convert.ToString(DtEstablo.Rows[0][0]), true);
            parameters[1] = new ReportParameter("MES", meses[fin.Month - 1] + " " + AñoParaElReporte.Year.ToString(), true);
            parameters[2] = new ReportParameter("DEC_UNO", dec1.ToString(), true);
            parameters[3] = new ReportParameter("DEC_DOS", dec2.ToString(), true);
            parameters[4] = new ReportParameter("DEC_TRES", dec3.ToString(), true);
            parameters[5] = new ReportParameter("NOVILLAS", _Novillas.ToString(), true);
            parameters[6] = new ReportParameter("VACAS", _Vacas.ToString(), true);
            parameters[7] = new ReportParameter("PRECIO_LECHE", "VENTA (PRECIO DE LA LECHE: " + precioBase.ToString("C3") + ")", true);
            parameters[8] = new ReportParameter("TITULONOID", tituloNoid, true);
            parameters[9] = new ReportParameter("VALIDACIONALIMENTO", validacionAlimento ? "si" : "no");
            parameters[10] = new ReportParameter("VALIDACIONMEDICINA", validacionMedicina ? "si" : "no");
            parameters[11] = new ReportParameter("VALIDACIONESVISIBLES", validacionesVisibles);

            LlenarListas(fecha_inicio, fecha_fin, horas);
            string mensaje = string.Empty;
            listaBacteriologia = controlador1.ListaLecheBacteriologia(fecha_inicio, fecha_fin, ref mensaje);
            datosProd = controlador1.DatosProduccion(fecha_inicio, fecha_fin, ref mensaje);
            DataTable dtHoja1Colorimetria = DTHoja1ConColorimetria(dtH1);
            DataTable dtHoja2Colirometria = DTHoja2ConColorimetria(dtH2);

            //try
            //{
            reportViewer1.LocalReport.DataSources.Clear();
            //ReportDataSource sour = new ReportDataSource("DataSet1", dtH1);
            ReportDataSource sour = new ReportDataSource("DataSet1", dtHoja1Colorimetria);
            reportViewer1.LocalReport.DataSources.Add(sour);
            //sour = new ReportDataSource("DataSet2", dtH2);
            sour = new ReportDataSource("DataSet2", dtHoja2Colirometria);
            reportViewer1.LocalReport.DataSources.Add(sour);
            sour = new ReportDataSource("DataSet3", _DtAux);
            reportViewer1.LocalReport.DataSources.Add(sour);
            sour = new ReportDataSource("DataSet4", _DtSR);
            reportViewer1.LocalReport.DataSources.Add(sour);
            reportViewer1.LocalReport.SetParameters(parameters);
            reportViewer1.LocalReport.Refresh();
            reportViewer1.RefreshReport();
            var deviceInfo = @"<DeviceInfo><EmbedFonts>Arial</EmbedFonts></DeviceInfo>";
            byte[] Bytes = reportViewer1.LocalReport.Render(format: "PDF", deviceInfo);

            ruta += "\\DIA.pdf";
            using (FileStream stream = new FileStream(ruta, FileMode.Create))
            {
                stream.Write(Bytes, 0, Bytes.Length);
            }

            if (!origen)
                Process.Start(ruta);
            //}
            //catch (ReportViewerException ex) { MessageBox.Show(ex.Message, "ERROR!", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            Cursor = Cursors.Default;
        }

        private void ValoresConsisitentesAlimentacion(DateTime FFinal)
        {
            bendiciones.ReportePeriodo sobrante = new bendiciones.ReportePeriodo();
            int hcorte = 0, horas;
            bool EAProd, MsProd, MsCrecimiento, MsDesarollo, MsVaquillas, MsSecas, MsReto, ValConsisitentes;
            double EAProdValor = 0, MsProdValor = 0, MsCrecimientosValor = 0, MsDesarolloValor = 0, MsVaquillasValor = 0, MsSecasValor = 0, MsRetoValor = 0;
            string query = "Select * from ConfiguraDiario";
            DataTable dtConfig;
            Hora_Corte(out horas, out hcorte);
            List<bendiciones.ReportePeriodo> listaIngredientes = new List<bendiciones.ReportePeriodo>();
            List<bendiciones.ReportePeriodo> listaForraje = new List<bendiciones.ReportePeriodo>();
            List<bendiciones.ReportePeriodo> listaAlimento = new List<bendiciones.ReportePeriodo>();
            List<bendiciones.ReportePeriodo> listaAgua = new List<bendiciones.ReportePeriodo>();
            bendiciones.IndicadorReportePeriodo indicadoresProduccion = new bendiciones.IndicadorReportePeriodo();
            bendiciones.IndicadorReportePeriodo indicadoresReto = new bendiciones.IndicadorReportePeriodo();
            bendiciones.IndicadorReportePeriodo indicadoresDesteteUno = new bendiciones.IndicadorReportePeriodo();
            bendiciones.IndicadorReportePeriodo indicadoresDesteteDos = new bendiciones.IndicadorReportePeriodo();
            bendiciones.IndicadorReportePeriodo indicadoresSecas = new bendiciones.IndicadorReportePeriodo();
            bendiciones.IndicadorReportePeriodo indicadoresVaquillas = new bendiciones.IndicadorReportePeriodo();
            conn.QueryMovsio(query, out dtConfig);
            if (!Convert.ToBoolean(dtConfig.Rows[0][0]))
            {
                FFinal = FFinal.AddDays(-1);
            }

            query = "Select * from ValConAlim WHERE fecha = #@FechaDia#";
            query = query.Replace("@FechaDia", FFinal.ToString("yyyy/MM/dd"));
            DataTable dtInfoValConAlim;
            conn.QueryMovsio(query, out dtInfoValConAlim);
            if (dtInfoValConAlim.Rows.Count == 0)
            {

                GTH.ReportePeriodo(ran_id.ToString(), horas, "22", FFinal.Date, FFinal.Date, out listaIngredientes, out indicadoresReto, out sobrante);
                GTH.ReportePeriodo(ran_id.ToString(), horas, "32", FFinal.Date, FFinal.Date, out listaIngredientes, out indicadoresDesteteUno, out sobrante);
                GTH.ReportePeriodo(ran_id.ToString(), horas, "33", FFinal.Date, FFinal.Date, out listaIngredientes, out indicadoresDesteteDos, out sobrante);
                GTH.ReportePeriodo(ran_id.ToString(), horas, "21", FFinal.Date, FFinal.Date, out listaIngredientes, out indicadoresSecas, out sobrante);
                GTH.ReportePeriodo(ran_id.ToString(), horas, "34", FFinal.Date, FFinal.Date, out listaIngredientes, out indicadoresVaquillas, out sobrante);
                GTH.ReportePeriodo(ran_id.ToString(), horas, "10,11,12,13", FFinal.Date, FFinal.Date, out listaIngredientes, out indicadoresProduccion, out sobrante);
                //Evaluacion de Ea Producción
                if (indicadoresProduccion.ANIMELES == 0)
                {
                    EAProd = true;
                    EAProdValor = -1;
                }
                else
                {
                    EAProd = (indicadoresProduccion.EA >= 1.1 && indicadoresProduccion.EA <= 1.8) ? true : false;
                    EAProdValor = indicadoresProduccion.EA;
                }
                //Evaluación de MS Producción
                if (indicadoresProduccion.ANIMELES == 0)
                {
                    MsProd = true;
                    MsProdValor = -1;
                }
                else
                {
                    MsProd = (indicadoresProduccion.MS >= 18 && indicadoresProduccion.MS <= 30) ? true : false;
                    MsProdValor = indicadoresProduccion.MS;
                }
                //Evaluación de MS Destetadas Uno
                if (indicadoresDesteteUno.ANIMELES == 0)
                {
                    MsCrecimiento = true;
                    MsCrecimientosValor = -1;
                }
                else
                {
                    MsCrecimiento = (indicadoresDesteteUno.MS >= 2 && indicadoresDesteteUno.MS <= 8) ? true : false;
                    MsCrecimientosValor = indicadoresDesteteUno.MS;
                }
                //Evaluación de Ms Destete Dos
                if (indicadoresDesteteDos.ANIMELES == 0)
                {
                    MsDesarollo = true;
                    MsDesarolloValor = -1;
                }
                else
                {
                    MsDesarollo = (indicadoresDesteteDos.MS >= 6 && indicadoresDesteteDos.MS <= 11) ? true : false;
                    MsDesarolloValor = indicadoresDesteteDos.MS;
                }
                //Evaluación de Ms Vaquillas
                if (indicadoresVaquillas.ANIMELES == 0)
                {
                    MsVaquillas = true;
                    MsVaquillasValor = -1;
                }
                else
                {
                    MsVaquillas = (indicadoresVaquillas.MS >= 8 && indicadoresVaquillas.MS <= 15) ? true : false;
                    MsVaquillasValor = indicadoresVaquillas.MS;
                }
                //Evaluación de MS Secas
                if (indicadoresSecas.ANIMELES == 0)
                {
                    MsSecas = true;
                    MsSecasValor = -1;
                }
                else
                {
                    MsSecas = (indicadoresSecas.MS >= 10 && indicadoresSecas.MS <= 19) ? true : false;
                    MsSecasValor = indicadoresSecas.MS;
                }
                //Evaluación de MS Reto
                if (indicadoresReto.ANIMELES == 0)
                {
                    MsReto = true;
                    MsRetoValor = -1;
                }
                else
                {
                    MsReto = (indicadoresReto.MS >= 9 && indicadoresReto.MS <= 19) ? true : false;
                    MsRetoValor = indicadoresReto.MS;

                }
                //Evaluación de todas la etapas
                if (EAProd && MsProd && MsCrecimiento && MsDesarollo && MsVaquillas && MsSecas && MsReto)
                {
                    ValConsisitentes = true;
                }
                else
                {
                    ValConsisitentes = false;
                }

                query = "Insert into ValConAlim values (@Id, #@Fecha#, @EAProd, @MsProd, @MsCrecimiento, @MsDesarollo, @MsVaquillas, @MsSecas, @MsReto, @ValConsisitentes, @ValEAProd, @ValMsProd, @ValMsCrecimiento, @ValMsDesarollo, @ValMsVaquillas, @ValMsSecas, @ValMsReto)";
                query = query.Replace("@Id", ran_id.ToString())
                             .Replace("@Fecha", FFinal.ToString("yyyy/MM/dd"))
                             .Replace("@EAProd", EAProd.ToString())
                             .Replace("@MsProd", MsProd.ToString())
                             .Replace("@MsCrecimiento", MsCrecimiento.ToString())
                             .Replace("@MsDesarollo", MsDesarollo.ToString())
                             .Replace("@MsVaquillas", MsVaquillas.ToString())
                             .Replace("@MsSecas", MsSecas.ToString())
                             .Replace("@MsReto", MsReto.ToString())
                             .Replace("@ValConsisitentes", ValConsisitentes.ToString())
                             .Replace("@ValEAProd", EAProdValor.ToString())
                             .Replace("@ValMsProd", MsProdValor.ToString())
                             .Replace("@ValMsCrecimiento", MsCrecimientosValor.ToString())
                             .Replace("@ValMsDesarollo", MsDesarolloValor.ToString())
                             .Replace("@ValMsVaquillas", MsVaquillasValor.ToString())
                             .Replace("@ValMsSecas", MsSecasValor.ToString())
                             .Replace("@ValMsReto", MsRetoValor.ToString());
                conn.QueryMovsio(query);

            }
            query = @"DELETE FROM ValConAlimPeriodo";
            conn.QueryMovsio(query);

            int promOrd = 0, promSecas = 0, promReto = 0, promDest1 = 0, promDest2 = 0, promVP = 0;
            for (int i = 0; i < 36; i++)
            {
                EAProdValor = MsProdValor = MsCrecimientosValor = MsDesarolloValor = MsVaquillasValor = MsSecasValor = MsRetoValor = 0;
                //GTH.ReportePeriodo(ran_id.ToString(), horas, "22", FFinal.AddDays(-i).Date, FFinal.AddDays(-i).Date, out listaIngredientes, out indicadoresReto, out sobrante);
                //GTH.ReportePeriodo(ran_id.ToString(), horas, "32", FFinal.AddDays(-i).Date, FFinal.AddDays(-i).Date, out listaIngredientes, out indicadoresDesteteUno, out sobrante);
                //GTH.ReportePeriodo(ran_id.ToString(), horas, "33", FFinal.AddDays(-i).Date, FFinal.AddDays(-i).Date, out listaIngredientes, out indicadoresDesteteDos, out sobrante);
                //GTH.ReportePeriodo(ran_id.ToString(), horas, "21", FFinal.AddDays(-i).Date, FFinal.AddDays(-i).Date, out listaIngredientes, out indicadoresSecas, out sobrante);
                //GTH.ReportePeriodo(ran_id.ToString(), horas, "34", FFinal.AddDays(-i).Date, FFinal.AddDays(-i).Date, out listaIngredientes, out indicadoresVaquillas, out sobrante);
                //GTH.ReportePeriodo(ran_id.ToString(), horas, "10,11,12,13", FFinal.AddDays(-i).Date, FFinal.AddDays(-i).Date, out listaIngredientes, out indicadoresProduccion, out sobrante);
                query = @"SELECT ran_id 
	                             ,ia_fecha                                                          AS Fecha
                                 ,IIF(ia_destetadas IS NOT NULL,ia_destetadas ,0)                   AS DesteteUno
                                 ,IIF(ia_destetadas2 IS NOT NULL,ia_destetadas2,0)                  AS DesteteDos
                                 ,IIF(ia_vaquillas IS NOT NULL,ia_vaquillas,0)                      AS Vaquillas
                                 ,IIF(ia_vacas_secas IS NOT NULL,ia_vacas_secas,0)                  AS Secas
                                 ,IIF(ia_vacas_ord IS NOT NULL,ia_vacas_ord,0)                      AS Ordeño
                                 ,IIF((ia_vqreto + ia_vcreto) IS NOT NULL ,ia_vqreto + ia_vcreto,0) AS Reto
                          FROM inventario_afi 
                          WHERE ia_fecha = '@FechaFinal'
                          AND ran_id = @Ran_ID";
                query = query.Replace("@FechaFinal", FFinal.AddDays(-i).Date.ToString("yyyy/MM/dd"))
                             .Replace("@Ran_ID", ran_id.ToString());
                DataTable dtAnimalesEtapas = new DataTable();
                conn.QueryAlimento(query, out dtAnimalesEtapas);

                query = @"SELECT  ROUND(AVG(d1))  AS Dest1
                               ,ROUND(AVG(d2))  AS Dest2
                               ,ROUND(AVG(v))  AS VP
                               ,ROUND(AVG(s)) AS SECAS
                               ,ROUND(AVG(o)) AS ORD
                               ,ROUND(AVG(r)) AS RETO
                        FROM
                        (
	                        SELECT  IIF(destetadas IS NOT NULL,destetadas ,0)             AS d1
	                               ,IIF(destetadas2 IS NOT NULL,destetadas2,0)            AS d2
	                               ,IIF(vaquillas IS NOT NULL,vaquillas,0)                AS v
	                               ,IIF(vacassecas IS NOT NULL,vacassecas,0)              AS s
	                               ,IIF(vacasordeña IS NOT NULL,vacasordeña,0)            AS o
	                               ,IIF((vqreto + vcreto) IS NOT NULL ,vqreto + vcreto,0) AS r
	                        FROM INVENTARIOAFIXDIA i
	                        WHERE FECHA BETWEEN @fechaI AND @fechaF
                        ) Tabla";
                query = query.Replace("@fechaI", ConvertToJulian(FFinal.AddDays(-i - 4)).ToString()).Replace("@fechaF", ConvertToJulian(FFinal.AddDays(-i - 1)).ToString());
                DataTable dtAnimalesPromedios = new DataTable();
                conn.QueryMovsio(query, out dtAnimalesPromedios);

                if (dtAnimalesPromedios.Rows.Count > 0)
                {
                    Int32.TryParse(dtAnimalesPromedios.Rows[0]["ORD"].ToString(), out promOrd);
                    Int32.TryParse(dtAnimalesPromedios.Rows[0]["SECAS"].ToString(), out promSecas);
                    Int32.TryParse(dtAnimalesPromedios.Rows[0]["RETO"].ToString(), out promReto);
                    Int32.TryParse(dtAnimalesPromedios.Rows[0]["Dest1"].ToString(), out promDest1);
                    Int32.TryParse(dtAnimalesPromedios.Rows[0]["Dest2"].ToString(), out promDest2);
                    Int32.TryParse(dtAnimalesPromedios.Rows[0]["VP"].ToString(), out promVP);
                }
                else
                {
                    promOrd = 0; promSecas = 0; promReto = 0; promDest1 = 0; promDest2 = 0; promVP = 0;
                }

                /*
                query = @" SELECT AVG(EA)
                                 ,AVG(MS)
                                 ,AVG(MS_DI * MH_DI)
                                 ,AVG(MS_DII * MH_DII)
                                 ,AVG(MS_VQ * MH_VQ)
                                 ,AVG(MS_S * MH_S)
                                 ,AVG(MS_R * MH_R)
                          FROM MPRODUC
                          WHERE fecha BETWEEN " + ConvertToJulian(FFinal.AddDays(-i -4)) + " AND " + ConvertToJulian(FFinal.AddDays(-i -1));
                */
                query = @"SELECT  IIF(SUM(CONTEA) > 0,SUM(EA) / SUM(CONTEA),0)    AS EA2
                                ,IIF(SUM(CONTMS) > 0,SUM(MS) / SUM(CONTMS),0)    AS MS2
                                ,IIF(SUM(CONTDI) > 0,SUM(DI) / SUM(CONTDI),0)    AS DI2
                                ,IIF(SUM(CONTDII) > 0,SUM(DII) / SUM(CONTDII),0) AS DII2
                                ,IIF(SUM(CONTVQ) > 0,SUM(VQ) / SUM(CONTVQ),0)    AS VQ2
                                ,IIF(SUM(CONTS) > 0,SUM(S) / SUM(CONTS),0)       AS S2
                                ,IIF(SUM(CONTR) > 0,SUM(R) / SUM(CONTR),0)       AS R2
                        FROM
                        (
	                      SELECT  m.EA
                               ,IIF(m.EA > 0,1,0)                    AS CONTEA
                               ,m.MS
                               ,IIF(MS > 0,1,IIF(i.VACASORDEÑA > 0,1,0)) AS CONTMS
                               ,MS_DI * MH_DI                         AS DI
                               ,IIF((MS_DI * MH_DI) > 0,1,IIF(i.destetadas > 0,1,0))          AS CONTDI
                               ,MS_DII * MH_DII                       AS DII
                               ,IIF((MS_DII * MH_DII) > 0,1,IIF(i.destetadas2 > 0,1,0))        AS CONTDII
                               ,MS_VQ * MH_VQ                         AS VQ
                               ,IIF((MS_VQ * MH_VQ ) > 0,1,IIF(i.vaquillas > 0,1,0))         AS CONTVQ
                               ,MS_S * MH_S                           AS S
                               ,IIF((MS_S * MH_S) > 0,1,IIF(i.vacassecas > 0,1,0))            AS CONTS
                               ,MS_R * MH_R                           AS R
                               ,IIF((MS_R * MH_R) > 0,1,IIF((i.vqreto + i.vcreto ) > 0,1,0))            AS CONTR
                        FROM MPRODUC m
                        INNER JOIN INVENTARIOAFIXDIA i
                        ON m.FECHA = i.FECHA
                        WHERE m.FECHA BETWEEN " + ConvertToJulian(FFinal.AddDays(-i - 4)) + " AND " + ConvertToJulian(FFinal.AddDays(-i - 1))
                        + " ) Tabla";
                DataTable dtInfoValConAlimPeriodo;
                conn.QueryMovsio(query, out dtInfoValConAlimPeriodo);

                if (dtInfoValConAlimPeriodo.Rows.Count > 0)
                {
                    EAProdValor = Convert.ToDouble(dtInfoValConAlimPeriodo.Rows[0][0].ToString());
                    MsProdValor = Convert.ToDouble(dtInfoValConAlimPeriodo.Rows[0][1].ToString());
                    MsCrecimientosValor = Convert.ToDouble(dtInfoValConAlimPeriodo.Rows[0][2].ToString());
                    MsDesarolloValor = Convert.ToDouble(dtInfoValConAlimPeriodo.Rows[0][3].ToString());
                    MsVaquillasValor = Convert.ToDouble(dtInfoValConAlimPeriodo.Rows[0][4].ToString());
                    MsSecasValor = Convert.ToDouble(dtInfoValConAlimPeriodo.Rows[0][5].ToString());
                    MsRetoValor = Convert.ToDouble(dtInfoValConAlimPeriodo.Rows[0][6].ToString());
                }

                if (dtAnimalesEtapas.Rows.Count > 0)
                {
                    if (Convert.ToInt32(dtAnimalesEtapas.Rows[0]["Ordeño"]) == 0)
                    {
                        EAProd = true;
                        EAProdValor = 0;
                    }
                    else
                    {
                        EAProd = promOrd > 0 ? (EAProdValor >= 1.1 && EAProdValor <= 1.8) ? true : false : true;
                    }
                    //Evaluación de MS Producción
                    if (Convert.ToInt32(dtAnimalesEtapas.Rows[0]["Ordeño"]) == 0)
                    {
                        MsProd = true;
                        MsProdValor = 0;
                    }
                    else
                    {
                        MsProd = promOrd > 0 ? (MsProdValor >= 18 && MsProdValor <= 30) ? true : false : true;
                    }
                    //Evaluación de MS Destetadas Uno 
                    if (Convert.ToInt32(dtAnimalesEtapas.Rows[0]["DesteteUno"]) == 0)
                    {
                        MsCrecimiento = true;
                        MsCrecimientosValor = 0;
                    }
                    else
                    {
                        MsCrecimiento = promDest1 > 0 ? (MsCrecimientosValor >= 2 && MsCrecimientosValor <= 8) ? true : false : true;
                    }
                    //Evaluación de Ms Destete Dos
                    if (Convert.ToInt32(dtAnimalesEtapas.Rows[0]["DesteteDos"]) == 0)
                    {
                        MsDesarollo = true;
                        MsDesarolloValor = 0;
                    }
                    else
                    {
                        MsDesarollo = promDest2 > 0 ? (MsDesarolloValor >= 6 && MsDesarolloValor <= 11) ? true : false : true;
                    }
                    //Evaluación de Ms Vaquillas
                    if (Convert.ToInt32(dtAnimalesEtapas.Rows[0]["Vaquillas"]) == 0)
                    {
                        MsVaquillas = true;
                        MsVaquillasValor = 0;
                    }
                    else
                    {
                        MsVaquillas = promVP != 0 ? (MsVaquillasValor >= 8 && MsVaquillasValor <= 15) ? true : false : true;
                    }
                    //Evaluación de MS Secas
                    if (Convert.ToInt32(dtAnimalesEtapas.Rows[0]["Secas"]) == 0)
                    {
                        MsSecas = true;
                        MsSecasValor = 0;
                    }
                    else
                    {
                        MsSecas = promSecas > 0 ? (MsSecasValor >= 10 && MsSecasValor <= 19) ? true : false : true;
                    }
                    //Evaluación de MS Reto
                    if (Convert.ToInt32(dtAnimalesEtapas.Rows[0]["Reto"]) == 0)
                    {
                        MsReto = true;
                        MsRetoValor = 0;
                    }
                    else
                    {
                        MsReto = promReto > 0 ? (MsRetoValor >= 9 && MsRetoValor <= 19) ? true : false : true;

                    }
                    //Evaluación de todas la etapas
                    if (EAProd && MsProd && MsCrecimiento && MsDesarollo && MsVaquillas && MsSecas && MsReto)
                    {
                        ValConsisitentes = true;
                    }
                    else
                    {
                        ValConsisitentes = false;
                    }
                }
                else
                {
                    EAProd = (EAProdValor >= 1.1 && EAProdValor <= 1.8) ? true : false;
                    MsProd = (MsProdValor >= 18 && MsProdValor <= 30) ? true : false;
                    MsCrecimiento = (MsCrecimientosValor >= 2 && MsCrecimientosValor <= 8) ? true : false;
                    MsDesarollo = (MsDesarolloValor >= 6 && MsDesarolloValor <= 11) ? true : false;
                    MsVaquillas = (MsVaquillasValor >= 8 && MsVaquillasValor <= 15) ? true : false;
                    MsSecas = (MsSecasValor >= 10 && MsSecasValor <= 19) ? true : false;
                    MsReto = (MsRetoValor >= 9 && MsRetoValor <= 19) ? true : false;
                    if (EAProd && MsProd && MsCrecimiento && MsDesarollo && MsVaquillas && MsSecas && MsReto)
                        ValConsisitentes = true;
                    else
                        ValConsisitentes = false;
                }


                query = "Insert into ValConAlimPeriodo values (@Id, #@Fecha#, @EAProd, @MsProd, @MsCrecimiento, @MsDesarollo, @MsVaquillas, @MsSecas, @MsReto, @ValConsisitentes, @ValEAProd, @ValMsProd, @ValMsCrecimiento, @ValMsDesarollo, @ValMsVaquillas, @ValMsSecas, @ValMsReto)";
                query = query.Replace("@Id", ran_id.ToString())
                             .Replace("@Fecha", FFinal.AddDays(-i).ToString("yyyy/MM/dd"))
                             .Replace("@EAProd", EAProd.ToString())
                             .Replace("@MsProd", MsProd.ToString())
                             .Replace("@MsCrecimiento", MsCrecimiento.ToString())
                             .Replace("@MsDesarollo", MsDesarollo.ToString())
                             .Replace("@MsVaquillas", MsVaquillas.ToString())
                             .Replace("@MsSecas", MsSecas.ToString())
                             .Replace("@MsReto", MsReto.ToString())
                             .Replace("@ValConsisitentes", ValConsisitentes.ToString())
                             .Replace("@ValEAProd", EAProdValor.ToString())
                             .Replace("@ValMsProd", MsProdValor.ToString())
                             .Replace("@ValMsCrecimiento", MsCrecimientosValor.ToString())
                             .Replace("@ValMsDesarollo", MsDesarolloValor.ToString())
                             .Replace("@ValMsVaquillas", MsVaquillasValor.ToString())
                             .Replace("@ValMsSecas", MsSecasValor.ToString())
                             .Replace("@ValMsReto", MsRetoValor.ToString());
                conn.QueryMovsio(query);

            }


            if (!Convert.ToBoolean(dtConfig.Rows[0][0]))
            {
                FFinal = FFinal.AddDays(1);
            }

        }
        private void Hoja1(int julianaI, int julianaF, bool noIdReal, out DataTable dt)
        {
            //if (ran_id != 33 && ran_id != 39)// si es covadonga o cañada nos quedamos con las fechas dadas por el calendario 
            //{
            HoraCorte(out fecha_inicio, out fecha_fin);
            //}
            ColumnasDTH1(out dt);
            DataTable dtA;
            #region query 
            /*
            string query = @"SELECT  CDATE(T2.FECHA)
                                       ,IIF(T2.TOTAL = NULL, 0 ,T2.TOTAL)
                                       ,T2.ORDEÑO
                                       ,T2.SECAS
                                       ,T2.HATO
                                       ,T2.PLACT
                                       ,T2.PPROT
                                       ,T2.UREA
                                       ,T2.GRASA
                                       ,T2.CCS
                                       ,T2.CTD
                                       ,T2.LECPROD
                                       ,T2.ANTIB
                                       ,T2.X
                                       ,T2.TOTALP
                                       ,T2.DELORD
                                       ,T2.VACAANTIB
                                       ,p.PRODUCCION
                                       ,T2.ILCA
                                       ,T2.IC
                                       ,T1.EA
                                       ,T1.ILCA_P
                                       ,T2.IC_P
                                       ,T1.L
                                       ,T2.MH
                                       ,T2.PMS
                                       ,T2.MS
                                       ,T2.SA
                                       ,T2.MSS
                                       ,T2.EAS
                                       ,T2.S
                                       ,T2.COSTO
                                       ,T2.PRECMS
                                       ,p.MS
                                       ,CRIBA.criba1
                                       ,CRIBA.criba2
                                       ,CRIBA.criba3
                                       ,CRIBA.criba4
                                       ,T2.NOIDSES1
                                       ,T2.NOIDSES2
                                       ,T2.NOIDSES3
                                FROM
                                (
                                    SELECT  T1.FECHA
                                           , T1.TOTAL
                                           , T1.ORDEÑO
                                           , T1.SECAS
                                           , T1.HATO
                                           , T1.PLACT
                                           , T1.PPROT
                                           , T1.UREA
                                           , T1.GRASA
                                           , T1.CCS
                                           , T1.CTD
                                           , T1.LECPROD
                                           , T1.ANTIB
                                           , T1.X
                                           , T1.TOTALP
                                           , T1.DELORD
                                           , T1.VACAANTIB
                                           , T1.ILCA
                                           , T1.IC
                                           , T1.EA
                                           , T1.ILCA_P
                                           , T1.IC_P
                                           , T1.L
                                           , T1.MH
                                           , T1.PMS
                                           , T1.MS
                                           , T1.SA
                                           , T1.MSS
                                           , T1.EAS
                                           , T1.S
                                           , T1.COSTO
                                           , T1.PRECMS
                                           , CRIBA.criba1
                                           , CRIBA.criba2
                                           , CRIBA.criba3
                                           , CRIBA.criba4
                                           , T1.NOIDSES1
                                           , T1.NOIDSES2
                                           , T1.NOIDSES3
                                    FROM
                                    (
                                        SELECT  T.FECHA
                                               , T.TOTAL
                                               , T.ORDEÑO
                                               , T.SECAS
                                               , T.HATO
                                               , T.PLACT
                                               , T.PPROT
                                               , T.UREA
                                               , T.GRASA
                                               , T.CCS
                                               , T.CTD
                                               , T.LECPROD
                                               , T.ANTIB
                                               , T.X
                                               , T.TOTALP
                                               , i.delord
                                               , T.VACAANTIB
                                               , T.ILCA
                                               , T.IC
                                               , T.EA
                                               , T.ILCA_P
                                               , T.IC_P
                                               , IIF(T.X > 0, T.COSTO / T.X, 0)      AS L
                                               , T.MH
                                               , IIF(T.MH > 0, T.MS / T.MH * 100, 0) AS PMS
                                               , T.MS
                                               , T.SA
                                               , T.MSS
                                               , T.EAS
                                               , IIF(T.MH > 0, T.SA / T.MH * 100, 0) AS S
                                               , T.COSTO
                                               , IIF(T.MS > 0, T.COSTO / T.MS, 0)    AS PRECMS
                                               , T.NOIDSES1
                                               , T.NOIDSES2
                                               , T.NOIDSES3
                                        FROM
                                        (
                                            SELECT  m.FECHA
                                                   , (LECFEDERAL + LECPLANTA)                                                AS TOTAL
                                                   , VACASORDEÑA                                                             AS ORDEÑO
                                                   , VACASSECAS                                                              AS SECAS
                                                   , VACASHATO                                                               AS HATO
                                                   , LACTOSA                                                                 AS PLACT
                                                   , PROTEINA                                                                AS PPROT
                                                   , d.UREA1                                                                 AS UREA
                                                   , d.Grasa1                                                                AS GRASA
                                                   , d.CCS1                                                                  AS CCS
                                                   , d.CTD1                                                                  AS CTD
                                                   , M.LECPROD
                                                   , M.ANTPROD                                                               AS ANTIB
                                                   , IIF(vacasordeña > 0, ROUND((lecprod + antprod) / vacasordeña, 2), 0)    AS X
                                                   , (M.LECPROD + M.ANTPROD)                                                 AS TOTALP
                                                   , M.VACAANTIB
                                                   , M.ILCA
                                                   , M.IC
                                                   , M.EA
                                                   , M.ILCA_P
                                                   , M.IC_P
                                                   , M.COSTO
                                                   , M.MH
                                                   , M.MS
                                                   , M.SA
                                                   , M.MSS
                                                   , M.EAS
                                                   , M.NOIDSES1
                                                   , M.NOIDSES2
                                                   , M.NOIDSES3
                                            FROM MPRODUC M

                                            LEFT JOIN
                                            (
                                                SELECT  RESULTADOS.FECHA
                                                       , AVG(RESULTADOS.PROTES) AS Proteina1
                                                       , SUM(RESULTADOS.UREAS)  AS Urea1
                                                       , SUM(RESULTADOS.GRASAS) AS Grasa1
                                                       , SUM(RESULTADOS.CCSS)   AS CCS1
                                                       , SUM(RESULTADOS.CTDS)   AS CTD1
                                                FROM
                                                (
                                                    SELECT  LECHEXDIA.FECHA
                                                           , VALORES.PROTEINA                                                                        AS PROTES
                                                           , IIF(ISNULL(LECHEXDIA.LG), NULL, (VALORES.LITROSXTANQUE / LECHEXDIA.LG) * VALORES.GRASA) AS GRASAS
                                                           , IIF(ISNULL(LECHEXDIA.LU), NULL, (VALORES.LITROSXTANQUE / LECHEXDIA.LU) * VALORES.UREA)  AS UREAS
                                                           , IIF(ISNULL(LECHEXDIA.LCC), NULL, (VALORES.LITROSXTANQUE / LECHEXDIA.LCC) * VALORES.CCS) AS CCSS
                                                           , IIF(ISNULL(LECHEXDIA.LCT), NULL, (VALORES.LITROSXTANQUE / LECHEXDIA.LCT) * VALORES.CTD) AS CTDS
                                                    FROM
                                                    ( 
                                                        SELECT  LITROSxDIA.DIA              AS FECHA
                                                               , SUM(LITROSxDIA.LITROGRASA) AS LG
                                                               , SUM(LITROSxDIA.LITROUREA)  AS LU
                                                               , SUM(LITROSxDIA.LITROCCS)   AS LCC
                                                               , SUM(LITROSxDIA.LITROCTS)   AS LCT
                                                        FROM
                                                        (
                                                            SELECT  FECHA                        AS DIA
                                                                   , GRASA
                                                                   , IIF(GRASA > 0, LITROSXTANQUE) AS LITROGRASA
                                                                   , IIF(UREA > 0, LITROSXTANQUE)  AS LITROUREA
                                                                   , IIF(CCS > 0, LITROSXTANQUE)   AS LITROCCS
                                                                   , IIF(CTD > 0, LITROSXTANQUE)   AS LITROCTS
                                                            FROM dproduc
                                                            WHERE FECHA BETWEEN  @julianaI  AND  @julianaF
                                                            ORDER BY FECHA
                                                        ) LITROSxDIA
                                                        GROUP BY  LITROSxDIA.DIA
                                                    ) LECHEXDIA
                                                    LEFT JOIN
                                                    (
                                                        SELECT  FECHA
                                                               , PROTEINA
                                                               , LITROSXTANQUE
                                                               , GRASA
                                                               , UREA
                                                               , CCS
                                                               , CTD
                                                        FROM dproduc
                                                        WHERE FECHA BETWEEN  @julianaI  AND  @julianaF
                                                        ORDER BY FECHA
                                                    ) VALORES
                                                    ON VALORES.FECHA = LECHEXDIA.FECHA
                                                )RESULTADOS
                                                GROUP BY  RESULTADOS.FECHA
                                            ) d
                                            ON m.FECHA = d.FECHA
                                            WHERE m.FECHA BETWEEN  @julianaI  AND  @julianaF
                                            ORDER BY m.FECHA
                                        ) T
                                        LEFT JOIN inventario i
                                        ON i.FECHA = T.FECHA
                                        ORDER BY 1
                                    ) T1
                                    LEFT JOIN
                                    (
                                        SELECT  FECHA
                                               , IIF(ISNULL(AVG(nivel1)), 0, AVG(nivel1))       AS criba1
                                               , IIF(ISNULL(AVG(nivel2)), 0, AVG(nivel2))       AS criba2
                                               , IIF(ISNULL(AVG(nivel3)), 0, AVG(nivel3))       AS criba3
                                               , IIF(ISNULL(AVG(nivel4)), 0, AVG(nivel4))       AS criba4
                                        FROM NIVELCRIBA
                                        WHERE FECHA BETWEEN  @julianaI  AND  @julianaF
                                        GROUP BY  FECHA
                                    )CRIBA
                                    ON T1.FECHA = CRIBA.FECHA
                                ) T2
                                LEFT JOIN PRECIOSTEORICOS p
                                ON T2.FECHA = p.FECHA 
                                order by T2.FECHA";
            */
            #endregion
            string query = QueryHoja1(noIdReal);
            query = query.Replace("@julianaI", julianaI.ToString())
                         .Replace("@julianaF", julianaF.ToString());
            conn.QueryMovsio(query, out dtA);

            DateTime dia;
            double totalT = 0, ordT = 0, secT = 0, hatoT = 0;
            DataTable DtPrecioLeche = new DataTable();
            for (int i = 0; i < dtA.Rows.Count; i++)
            {
                DataRow row = dt.NewRow();
                dia = Convert.ToDateTime(dtA.Rows[i][0]);
                row[0] = dia.Day.ToString();
                for (int j = 1; j < dtA.Columns.Count; j++)
                    row[j] = dtA.Rows[i][j];

                dt.Rows.Add(row);

            }
            AddRow(dt);
            int[] posiciones = { 1, 2, 3, 4, 11, 12, 14, 38, 39, 40 };
            Total(dt, posiciones);

            PromediosH1(dt, noIdReal);
            DiferenciasH1(dt);
            ArrayList columnasH1 = new ArrayList();
            columnasH1.Add(17);
            columnasH1.Add(33);
            RemoverCeros(dt, columnasH1);
            dt.Rows[32][34] = 0;
            dt.Rows[32][35] = 0;
            dt.Rows[32][36] = 0;
            dt.Rows[32][37] = 0;
            //En esta parte tomamos valores de data table para obtener los promedios de N1, N2, N3 y N4
            for (int i = 34; i < 38; i++)
            {
                int _contador = 0;
                for (int J = 0; J <= 30; J++)
                {
                    if (dt.Rows[J][i] == DBNull.Value)
                    {

                    }
                    else
                    {
                        dt.Rows[32][i] = Convert.ToInt32(dt.Rows[J][i]) + Convert.ToInt32(dt.Rows[32][i]);
                        _contador++;
                    }
                }
                if (Convert.ToInt32(dt.Rows[32][i]) == 0)
                {

                }
                else
                {
                    dt.Rows[32][i] = Convert.ToInt32(dt.Rows[32][i]) / _contador;
                }
            }

            //Se dimenciona el datatable 
            DtPrecioLeche.Columns.Add("FECHA").DataType = System.Type.GetType("System.Int32");
            DtPrecioLeche.Columns.Add("PRECIO").DataType = System.Type.GetType("System.Int32");

            for (int i = 0; i < 31; i++)
            {
                DtPrecioLeche.Rows.Add();
            }
            // se buscan los valores del precio de la leche respecto a las fechas
            query = "SELECT DAY(CDATE(FECHA)), PRECIOLECHE FROM MPRODUC "
                     + " WHERE FECHA BETWEEN " + julianaI + " AND " + julianaF
                     + " ORDER BY FECHA";
            conn.QueryMovGanado(query, out DtPrecioLeche);
            PonderadosHoja1(dt, DtPrecioLeche);
        }

        private string QueryHoja1(bool noIdReal)
        {
            string query = "";

            if (noIdReal)
            {
                query = @"SELECT  CDATE(T2.FECHA)
                                       ,IIF(T2.TOTAL = NULL, 0 ,T2.TOTAL)
                                       ,T2.ORDEÑO
                                       ,T2.SECAS
                                       ,T2.HATO
                                       ,T2.PLACT
                                       ,T2.PPROT
                                       ,T2.UREA
                                       ,T2.GRASA
                                       ,T2.CCS
                                       ,T2.CTD
                                       ,T2.LECPROD
                                       ,T2.ANTIB
                                       ,T2.X
                                       ,T2.TOTALP
                                       ,T2.DELORD
                                       ,T2.VACAANTIB
                                       ,p.PRODUCCION
                                       ,T2.ILCA
                                       ,T2.IC
                                       ,T1.EA
                                       ,T1.ILCA_P
                                       ,T2.IC_P
                                       ,T1.L
                                       ,T2.MH
                                       ,T2.PMS
                                       ,T2.MS
                                       ,T2.SA
                                       ,T2.MSS
                                       ,T2.EAS
                                       ,T2.S
                                       ,T2.COSTO
                                       ,T2.PRECMS
                                       ,p.MS
                                       ,CRIBA.criba1
                                       ,CRIBA.criba2
                                       ,CRIBA.criba3
                                       ,CRIBA.criba4
                                       ,T2.NOIDSES1
                                       ,T2.NOIDSES2
                                       ,T2.NOIDSES3
                                FROM
                                (
                                    SELECT  T1.FECHA
                                           , T1.TOTAL
                                           , T1.ORDEÑO
                                           , T1.SECAS
                                           , T1.HATO
                                           , T1.PLACT
                                           , T1.PPROT
                                           , T1.UREA
                                           , T1.GRASA
                                           , T1.CCS
                                           , T1.CTD
                                           , T1.LECPROD
                                           , T1.ANTIB
                                           , T1.X
                                           , T1.TOTALP
                                           , T1.DELORD
                                           , T1.VACAANTIB
                                           , T1.ILCA
                                           , T1.IC
                                           , T1.EA
                                           , T1.ILCA_P
                                           , T1.IC_P
                                           , T1.L
                                           , T1.MH
                                           , T1.PMS
                                           , T1.MS
                                           , T1.SA
                                           , T1.MSS
                                           , T1.EAS
                                           , T1.S
                                           , T1.COSTO
                                           , T1.PRECMS
                                           , CRIBA.criba1
                                           , CRIBA.criba2
                                           , CRIBA.criba3
                                           , CRIBA.criba4
                                           , T1.NOIDSES1
                                           , T1.NOIDSES2
                                           , T1.NOIDSES3
                                    FROM
                                    (
                                        SELECT  T.FECHA
                                               , T.TOTAL
                                               , T.ORDEÑO
                                               , T.SECAS
                                               , T.HATO
                                               , T.PLACT
                                               , T.PPROT
                                               , T.UREA
                                               , T.GRASA
                                               , T.CCS
                                               , T.CTD
                                               , T.LECPROD
                                               , T.ANTIB
                                               , T.X
                                               , T.TOTALP
                                               , i.delord
                                               , T.VACAANTIB
                                               , T.ILCA
                                               , T.IC
                                               , T.EA
                                               , T.ILCA_P
                                               , T.IC_P
                                               , IIF(T.X > 0, T.COSTO / T.X, 0)      AS L
                                               , T.MH
                                               , IIF(T.MH > 0, T.MS / T.MH * 100, 0) AS PMS
                                               , T.MS
                                               , T.SA
                                               , T.MSS
                                               , T.EAS
                                               , IIF(T.MH > 0, T.SA / T.MH * 100, 0) AS S
                                               , T.COSTO
                                               , IIF(T.MS > 0, T.COSTO / T.MS, 0)    AS PRECMS
                                               , T.NOIDSES1
                                               , T.NOIDSES2
                                               , T.NOIDSES3
                                        FROM
                                        (
                                            SELECT  m.FECHA
                                                   , (LECFEDERAL + LECPLANTA)                                                AS TOTAL
                                                   , VACASORDEÑA                                                             AS ORDEÑO
                                                   , VACASSECAS                                                              AS SECAS
                                                   , VACASHATO                                                               AS HATO
                                                   , LACTOSA                                                                 AS PLACT
                                                   , PROTEINA                                                                AS PPROT
                                                   , d.UREA1                                                                 AS UREA
                                                   , d.Grasa1                                                                AS GRASA
                                                   , d.CCS1                                                                  AS CCS
                                                   , d.CTD1                                                                  AS CTD
                                                   , M.LECPROD
                                                   , M.ANTPROD                                                               AS ANTIB
                                                   , IIF(vacasordeña > 0, ROUND((lecprod + antprod) / vacasordeña, 2), 0)    AS X
                                                   , (M.LECPROD + M.ANTPROD)                                                 AS TOTALP
                                                   , M.VACAANTIB
                                                   , M.ILCA
                                                   , M.IC
                                                   , M.EA
                                                   , M.ILCA_P
                                                   , M.IC_P
                                                   , M.COSTO
                                                   , M.MH
                                                   , M.MS
                                                   , M.SA
                                                   , M.MSS
                                                   , M.EAS
                                                   ,M.NOIDSES1REAL AS NOIDSES1
			                                       ,M.NOIDSES2REAL AS NOIDSES2
			                                       ,M.NOIDSES3REAL AS NOIDSES3
                                            FROM MPRODUC M
                                            LEFT JOIN
                                            (
                                                SELECT  RESULTADOS.FECHA
                                                       , AVG(RESULTADOS.PROTES) AS Proteina1
                                                       , SUM(RESULTADOS.UREAS)  AS Urea1
                                                       , SUM(RESULTADOS.GRASAS) AS Grasa1
                                                       , SUM(RESULTADOS.CCSS)   AS CCS1
                                                       , SUM(RESULTADOS.CTDS)   AS CTD1
                                                FROM
                                                (
                                                    SELECT  LECHEXDIA.FECHA
                                                           , VALORES.PROTEINA                                                                        AS PROTES
                                                           , IIF(ISNULL(LECHEXDIA.LG), NULL, (VALORES.LITROSXTANQUE / LECHEXDIA.LG) * VALORES.GRASA) AS GRASAS
                                                           , IIF(ISNULL(LECHEXDIA.LU), NULL, (VALORES.LITROSXTANQUE / LECHEXDIA.LU) * VALORES.UREA)  AS UREAS
                                                           , IIF(ISNULL(LECHEXDIA.LCC), NULL, (VALORES.LITROSXTANQUE / LECHEXDIA.LCC) * VALORES.CCS) AS CCSS
                                                           , IIF(ISNULL(LECHEXDIA.LCT), NULL, (VALORES.LITROSXTANQUE / LECHEXDIA.LCT) * VALORES.CTD) AS CTDS
                                                    FROM
                                                    ( 
                                                        SELECT  LITROSxDIA.DIA              AS FECHA
                                                               , SUM(LITROSxDIA.LITROGRASA) AS LG
                                                               , SUM(LITROSxDIA.LITROUREA)  AS LU
                                                               , SUM(LITROSxDIA.LITROCCS)   AS LCC
                                                               , SUM(LITROSxDIA.LITROCTS)   AS LCT
                                                        FROM
                                                        (
                                                            SELECT  FECHA                        AS DIA
                                                                   , GRASA
                                                                   , IIF(GRASA > 0, LITROSXTANQUE) AS LITROGRASA
                                                                   , IIF(UREA > 0, LITROSXTANQUE)  AS LITROUREA
                                                                   , IIF(CCS > 0, LITROSXTANQUE)   AS LITROCCS
                                                                   , IIF(CTD > 0, LITROSXTANQUE)   AS LITROCTS
                                                            FROM dproduc
                                                            WHERE FECHA BETWEEN  @julianaI  AND  @julianaF
                                                            ORDER BY FECHA
                                                        ) LITROSxDIA
                                                        GROUP BY  LITROSxDIA.DIA
                                                    ) LECHEXDIA
                                                    LEFT JOIN
                                                    (
                                                        SELECT  FECHA
                                                               , PROTEINA
                                                               , LITROSXTANQUE
                                                               , GRASA
                                                               , UREA
                                                               , CCS
                                                               , CTD
                                                        FROM dproduc
                                                        WHERE FECHA BETWEEN  @julianaI  AND  @julianaF
                                                        ORDER BY FECHA
                                                    ) VALORES
                                                    ON VALORES.FECHA = LECHEXDIA.FECHA
                                                )RESULTADOS
                                                GROUP BY  RESULTADOS.FECHA
                                            ) d
                                            ON m.FECHA = d.FECHA
                                            WHERE m.FECHA BETWEEN  @julianaI  AND  @julianaF
                                            ORDER BY m.FECHA
                                        ) T
                                        LEFT JOIN inventario i
                                        ON i.FECHA = T.FECHA
                                        ORDER BY 1
                                    ) T1
                                    LEFT JOIN
                                    (
                                        SELECT  FECHA
                                               , IIF(ISNULL(AVG(nivel1)), 0, AVG(nivel1))       AS criba1
                                               , IIF(ISNULL(AVG(nivel2)), 0, AVG(nivel2))       AS criba2
                                               , IIF(ISNULL(AVG(nivel3)), 0, AVG(nivel3))       AS criba3
                                               , IIF(ISNULL(AVG(nivel4)), 0, AVG(nivel4))       AS criba4
                                        FROM NIVELCRIBA
                                        WHERE FECHA BETWEEN  @julianaI  AND  @julianaF
                                        GROUP BY  FECHA
                                    )CRIBA
                                    ON T1.FECHA = CRIBA.FECHA
                                ) T2
                                LEFT JOIN PRECIOSTEORICOS p
                                ON T2.FECHA = p.FECHA 
                                order by T2.FECHA";
            }
            else
            {
                query = @"SELECT  CDATE(T2.FECHA)
                                       ,IIF(T2.TOTAL = NULL, 0 ,T2.TOTAL)
                                       ,T2.ORDEÑO
                                       ,T2.SECAS
                                       ,T2.HATO
                                       ,T2.PLACT
                                       ,T2.PPROT
                                       ,T2.UREA
                                       ,T2.GRASA
                                       ,T2.CCS
                                       ,T2.CTD
                                       ,T2.LECPROD
                                       ,T2.ANTIB
                                       ,T2.X
                                       ,T2.TOTALP
                                       ,T2.DELORD
                                       ,T2.VACAANTIB
                                       ,p.PRODUCCION
                                       ,T2.ILCA
                                       ,T2.IC
                                       ,T1.EA
                                       ,T1.ILCA_P
                                       ,T2.IC_P
                                       ,T1.L
                                       ,T2.MH
                                       ,T2.PMS
                                       ,T2.MS
                                       ,T2.SA
                                       ,T2.MSS
                                       ,T2.EAS
                                       ,T2.S
                                       ,T2.COSTO
                                       ,T2.PRECMS
                                       ,p.MS
                                       ,CRIBA.criba1
                                       ,CRIBA.criba2
                                       ,CRIBA.criba3
                                       ,CRIBA.criba4
                                       ,T2.NOIDSES1
                                       ,T2.NOIDSES2
                                       ,T2.NOIDSES3
                                FROM
                                (
                                    SELECT  T1.FECHA
                                           , T1.TOTAL
                                           , T1.ORDEÑO
                                           , T1.SECAS
                                           , T1.HATO
                                           , T1.PLACT
                                           , T1.PPROT
                                           , T1.UREA
                                           , T1.GRASA
                                           , T1.CCS
                                           , T1.CTD
                                           , T1.LECPROD
                                           , T1.ANTIB
                                           , T1.X
                                           , T1.TOTALP
                                           , T1.DELORD
                                           , T1.VACAANTIB
                                           , T1.ILCA
                                           , T1.IC
                                           , T1.EA
                                           , T1.ILCA_P
                                           , T1.IC_P
                                           , T1.L
                                           , T1.MH
                                           , T1.PMS
                                           , T1.MS
                                           , T1.SA
                                           , T1.MSS
                                           , T1.EAS
                                           , T1.S
                                           , T1.COSTO
                                           , T1.PRECMS
                                           , CRIBA.criba1
                                           , CRIBA.criba2
                                           , CRIBA.criba3
                                           , CRIBA.criba4
                                           , T1.NOIDSES1
                                           , T1.NOIDSES2
                                           , T1.NOIDSES3
                                    FROM
                                    (
                                        SELECT  T.FECHA
                                               , T.TOTAL
                                               , T.ORDEÑO
                                               , T.SECAS
                                               , T.HATO
                                               , T.PLACT
                                               , T.PPROT
                                               , T.UREA
                                               , T.GRASA
                                               , T.CCS
                                               , T.CTD
                                               , T.LECPROD
                                               , T.ANTIB
                                               , T.X
                                               , T.TOTALP
                                               , i.delord
                                               , T.VACAANTIB
                                               , T.ILCA
                                               , T.IC
                                               , T.EA
                                               , T.ILCA_P
                                               , T.IC_P
                                               , IIF(T.X > 0, T.COSTO / T.X, 0)      AS L
                                               , T.MH
                                               , IIF(T.MH > 0, T.MS / T.MH * 100, 0) AS PMS
                                               , T.MS
                                               , T.SA
                                               , T.MSS
                                               , T.EAS
                                               , IIF(T.MH > 0, T.SA / T.MH * 100, 0) AS S
                                               , T.COSTO
                                               , IIF(T.MS > 0, T.COSTO / T.MS, 0)    AS PRECMS
                                               , T.NOIDSES1
                                               , T.NOIDSES2
                                               , T.NOIDSES3
                                        FROM
                                        (
                                            SELECT  m.FECHA
                                                   , (LECFEDERAL + LECPLANTA)                                                AS TOTAL
                                                   , VACASORDEÑA                                                             AS ORDEÑO
                                                   , VACASSECAS                                                              AS SECAS
                                                   , VACASHATO                                                               AS HATO
                                                   , LACTOSA                                                                 AS PLACT
                                                   , PROTEINA                                                                AS PPROT
                                                   , d.UREA1                                                                 AS UREA
                                                   , d.Grasa1                                                                AS GRASA
                                                   , d.CCS1                                                                  AS CCS
                                                   , d.CTD1                                                                  AS CTD
                                                   , M.LECPROD
                                                   , M.ANTPROD                                                               AS ANTIB
                                                   , IIF(vacasordeña > 0, ROUND((lecprod + antprod) / vacasordeña, 2), 0)    AS X
                                                   , (M.LECPROD + M.ANTPROD)                                                 AS TOTALP
                                                   , M.VACAANTIB
                                                   , M.ILCA
                                                   , M.IC
                                                   , M.EA
                                                   , M.ILCA_P
                                                   , M.IC_P
                                                   , M.COSTO
                                                   , M.MH
                                                   , M.MS
                                                   , M.SA
                                                   , M.MSS
                                                   , M.EAS
                                                   , M.NOIDSES1
                                                   , M.NOIDSES2
                                                   , M.NOIDSES3
                                            FROM MPRODUC M

                                            LEFT JOIN
                                            (
                                                SELECT  RESULTADOS.FECHA
                                                       , AVG(RESULTADOS.PROTES) AS Proteina1
                                                       , SUM(RESULTADOS.UREAS)  AS Urea1
                                                       , SUM(RESULTADOS.GRASAS) AS Grasa1
                                                       , SUM(RESULTADOS.CCSS)   AS CCS1
                                                       , SUM(RESULTADOS.CTDS)   AS CTD1
                                                FROM
                                                (
                                                    SELECT  LECHEXDIA.FECHA
                                                           , VALORES.PROTEINA                                                                        AS PROTES
                                                           , IIF(ISNULL(LECHEXDIA.LG), NULL, (VALORES.LITROSXTANQUE / LECHEXDIA.LG) * VALORES.GRASA) AS GRASAS
                                                           , IIF(ISNULL(LECHEXDIA.LU), NULL, (VALORES.LITROSXTANQUE / LECHEXDIA.LU) * VALORES.UREA)  AS UREAS
                                                           , IIF(ISNULL(LECHEXDIA.LCC), NULL, (VALORES.LITROSXTANQUE / LECHEXDIA.LCC) * VALORES.CCS) AS CCSS
                                                           , IIF(ISNULL(LECHEXDIA.LCT), NULL, (VALORES.LITROSXTANQUE / LECHEXDIA.LCT) * VALORES.CTD) AS CTDS
                                                    FROM
                                                    ( 
                                                        SELECT  LITROSxDIA.DIA              AS FECHA
                                                               , SUM(LITROSxDIA.LITROGRASA) AS LG
                                                               , SUM(LITROSxDIA.LITROUREA)  AS LU
                                                               , SUM(LITROSxDIA.LITROCCS)   AS LCC
                                                               , SUM(LITROSxDIA.LITROCTS)   AS LCT
                                                        FROM
                                                        (
                                                            SELECT  FECHA                        AS DIA
                                                                   , GRASA
                                                                   , IIF(GRASA > 0, LITROSXTANQUE) AS LITROGRASA
                                                                   , IIF(UREA > 0, LITROSXTANQUE)  AS LITROUREA
                                                                   , IIF(CCS > 0, LITROSXTANQUE)   AS LITROCCS
                                                                   , IIF(CTD > 0, LITROSXTANQUE)   AS LITROCTS
                                                            FROM dproduc
                                                            WHERE FECHA BETWEEN  @julianaI  AND  @julianaF
                                                            ORDER BY FECHA
                                                        ) LITROSxDIA
                                                        GROUP BY  LITROSxDIA.DIA
                                                    ) LECHEXDIA
                                                    LEFT JOIN
                                                    (
                                                        SELECT  FECHA
                                                               , PROTEINA
                                                               , LITROSXTANQUE
                                                               , GRASA
                                                               , UREA
                                                               , CCS
                                                               , CTD
                                                        FROM dproduc
                                                        WHERE FECHA BETWEEN  @julianaI  AND  @julianaF
                                                        ORDER BY FECHA
                                                    ) VALORES
                                                    ON VALORES.FECHA = LECHEXDIA.FECHA
                                                )RESULTADOS
                                                GROUP BY  RESULTADOS.FECHA
                                            ) d
                                            ON m.FECHA = d.FECHA
                                            WHERE m.FECHA BETWEEN  @julianaI  AND  @julianaF
                                            ORDER BY m.FECHA
                                        ) T
                                        LEFT JOIN inventario i
                                        ON i.FECHA = T.FECHA
                                        ORDER BY 1
                                    ) T1
                                    LEFT JOIN
                                    (
                                        SELECT  FECHA
                                               , IIF(ISNULL(AVG(nivel1)), 0, AVG(nivel1))       AS criba1
                                               , IIF(ISNULL(AVG(nivel2)), 0, AVG(nivel2))       AS criba2
                                               , IIF(ISNULL(AVG(nivel3)), 0, AVG(nivel3))       AS criba3
                                               , IIF(ISNULL(AVG(nivel4)), 0, AVG(nivel4))       AS criba4
                                        FROM NIVELCRIBA
                                        WHERE FECHA BETWEEN  @julianaI  AND  @julianaF
                                        GROUP BY  FECHA
                                    )CRIBA
                                    ON T1.FECHA = CRIBA.FECHA
                                ) T2
                                LEFT JOIN PRECIOSTEORICOS p
                                ON T2.FECHA = p.FECHA 
                                order by T2.FECHA";
            }


            return query;
        }

        private void PromediosH1(DataTable dt, bool noIdReal)
        {
            DateTime inicio, fin;
            int julianaI = ConvertToJulian(fecha_inicio);
            int julianaF = ConvertToJulian(fecha_fin);
            /*  if (ran_id != 33 && ran_id != 39) //Se agregara para covadonga y cañada
              {*/
            HoraCorte(out inicio, out fin);
            DataTable dtI, dtV;
            Indicadores("10,11,12,13", "ia_vacas_ord", inicio, fin, fecha_inicio, fecha_fin, out dtI);
            VentaProduccion(julianaI, julianaF, noIdReal, out dtV);
            AddPromedioH1("PROM", dt, dtI, dtV);

            int julianaIAnt = ConvertToJulian(inicio.AddYears(-1));
            int julianaFAnt = ConvertToJulian(fin.AddYears(-1));
            Indicadores("10,11,12,13", "ia_vacas_ord", inicio.AddYears(-1), fin.AddYears(-1), fecha_inicio.AddYears(-1), fecha_fin.AddYears(-1), out dtI);
            VentaProduccion(julianaIAnt, julianaFAnt, noIdReal, out dtV);
            AddPromedioH1("AÑO ANT", dt, dtI, dtV);
            /*
                        }
                        else
                        {
                            DataRow row = dt.NewRow();
                            row[0] = "PROM";
                            dt.Rows.Add(row);
                            DataRow rows = dt.NewRow();
                            rows[0] = "AÑO ANT";
                            dt.Rows.Add(rows);
                        }
            */
        }

        private void AddPromedioH1(string titulo, DataTable dt, DataTable dtI, DataTable dtV)
        {
            DataRow row = dt.NewRow();
            row[0] = titulo;
            row[1] = dtV.Rows[0][0];
            row[2] = dtV.Rows[0][1];
            row[3] = dtV.Rows[0][2];
            row[4] = dtV.Rows[0][3];
            row[5] = dtV.Rows[0][4];
            row[6] = dtV.Rows[0][5];
            row[7] = dtV.Rows[0][6];
            row[8] = dtV.Rows[0][7];
            row[9] = dtV.Rows[0][8];
            row[10] = dtV.Rows[0][9];
            row[11] = dtV.Rows[0][10];
            row[12] = dtV.Rows[0][11];
            row[13] = dtI.Rows[0][1];
            row[14] = dtV.Rows[0][12];
            row[15] = dtV.Rows[0][13];
            row[16] = dtV.Rows[0][14];
            row[17] = DBNull.Value;
            row[18] = dtI.Rows[0][2];
            row[19] = dtI.Rows[0][3];
            row[20] = dtI.Rows[0][4];
            row[21] = dtI.Rows[0][5];
            row[22] = dtI.Rows[0][6];
            row[23] = dtI.Rows[0][7];
            row[24] = dtI.Rows[0][8];
            row[25] = dtI.Rows[0][9];
            row[26] = dtI.Rows[0][10];
            row[27] = dtI.Rows[0][11];
            row[28] = dtI.Rows[0][12];
            row[29] = dtI.Rows[0][13];
            row[30] = Convert.ToDouble(dtI.Rows[0][8]) > 0 ? Convert.ToDouble(dtI.Rows[0][11]) / Convert.ToDouble(dtI.Rows[0][8]) * 100 : 0;
            row[31] = dtI.Rows[0][14];
            row[32] = dtI.Rows[0][15];
            row[33] = DBNull.Value;
            row[34] = dtV.Rows[0][15];
            row[35] = dtV.Rows[0][16];
            row[36] = dtV.Rows[0][17];
            row[37] = dtV.Rows[0][18];
            row[38] = dtV.Rows[0][19];
            row[39] = dtV.Rows[0][20];
            row[40] = dtV.Rows[0][21];
            dt.Rows.Add(row);
        }

        private void DiferenciasH1(DataTable dt)
        {
            int ultimo = dt.Rows.Count - 1;
            int penultimo = ultimo - 1;
            double actual, anterior, diferencia, porcentaje;

            DataRow row = dt.NewRow();
            DataRow row2 = dt.NewRow();
            row[0] = "DIF %";
            row2[0] = "DIF #";
            for (int i = 1; i < dt.Columns.Count; i++)
            {
                if (i == 17 || i == 33)
                    continue;

                actual = dt.Rows[penultimo][i] != DBNull.Value ? Convert.ToDouble(dt.Rows[penultimo][i]) : 0;
                anterior = dt.Rows[ultimo][i] != DBNull.Value ? Convert.ToDouble(dt.Rows[ultimo][i]) : 0;
                diferencia = actual - anterior;
                porcentaje = actual != 0 ? diferencia / actual : 0;
                row[i] = porcentaje;
                row2[i] = diferencia;
            }
            dt.Rows.Add(row);
            dt.Rows.Add(row2);

        }

        private void RemoverCeros(DataTable dt, ArrayList posiciones)
        {
            int index;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 1; j < dt.Columns.Count; j++)
                {
                    index = posiciones.IndexOf(j);
                    if (index == -1)
                        dt.Rows[i][j] = dt.Rows[i][j] != DBNull.Value && Convert.ToDouble(dt.Rows[i][j]) == 0 ? DBNull.Value : dt.Rows[i][j];


                }
            }
        }

        private void VentaProduccion(int inicio, int fin, bool noIdReal, out DataTable dt)
        {
            string query = @"";

            if (noIdReal)
            {
                query = @"SELECT  IIF(ISNULL(AVG(T3.TOTAL)),NULL,CLng(AVG(T3.TOTAL)))
                                    ,IIF(ISNULL(AVG(T3.ORDEÑO)),NULL,CINT(AVG(T3.ORDEÑO)))                        AS ORDEÑO
                                    ,IIF(ISNULL(AVG(T3.SECAS)),NULL,CINT(AVG(T3.SECAS)))                          AS SECAS
                                    ,IIF(ISNULL(AVG(T3.HATO)),NULL,CINT(AVG(T3.HATO)))                            AS HATO
                                    ,AVG(T3.PLACT)                                                                AS LACT
                                    ,AVG(T3.PPROT)                                                                AS PROT
                                    ,AVG(T3.UREA)                                                                 AS UREA
                                    ,AVG(T3.GRASA)                                                                AS GRASA
                                    ,AVG(T3.CCS)                                                                  AS CCS
                                    ,IIF(ISNULL(AVG(T3.CTD)),NULL,CInt(AVG(T3.CTD)))                              AS CTD
                                    ,IIF(ISNULL(AVG(T3.LECPROD)),NULL,CLng(AVG(T3.LECPROD)))                      AS LECHE
                                    ,IIF(ISNULL(AVG(T3.ANTIB)),NULL,CINT(AVG(T3.ANTIB)))                          AS ANTIB
                                    ,IIF(ISNULL(AVG(T3.TOTALP)),NULL,CLNG(AVG(T3.TOTALP)))                        AS TOTAL2
                                    ,IIF(ISNULL(AVG(T3.DELORD)),NULL,CINT(AVG(T3.DELORD)))                        AS DEL
                                    ,IIF(ISNULL(AVG(T3.VACAANTIB)),NULL,CINT(AVG(T3.VACAANTIB)))                  AS ANT
                                    ,AVG(T3.criba1)                                                               AS N1
                                    ,AVG(T3.criba2)                                                               AS N2
                                    ,AVG(T3.criba3)                                                               AS N3
                                    ,AVG(T3.criba4)                                                               AS N4
                                    ,IIF(SUM(T3.ContNOIDSES1) > 0,CINT(SUM(T3.NOIDSES1)/SUM(T3.ContNOIDSES1)) ,0) AS SES1
                                    ,IIF(SUM(T3.ContNOIDSES2) > 0,CINT(SUM(T3.NOIDSES2)/SUM(T3.ContNOIDSES2)) ,0) AS SES2
                                    ,IIF(SUM(T3.ContNOIDSES3) > 0,CINT(SUM(T3.NOIDSES3)/SUM(T3.ContNOIDSES3)) ,0) AS SES3
                            FROM
                            (
	                            SELECT  T2.FECHA
	                                    ,T2.TOTAL
	                                    ,T2.ORDEÑO
	                                    ,T2.SECAS
	                                    ,T2.HATO
	                                    ,T2.PLACT
	                                    ,T2.PPROT
	                                    ,T2.UREA
	                                    ,T2.GRASA
	                                    ,T2.CCS
	                                    ,T2.CTD
	                                    ,T2.LECPROD
	                                    ,T2.ANTIB
	                                    ,T2.X
	                                    ,T2.TOTALP
	                                    ,T2.DELORD
	                                    ,T2.VACAANTIB
	                                    ,CRIBA.criba1
	                                    ,CRIBA.criba2
	                                    ,CRIBA.criba3
	                                    ,CRIBA.criba4
	                                    ,T2.NOIDSES1
	                                    ,T2.NOIDSES2
	                                    ,T2.NOIDSES3
	                                    ,T2.ContNOIDSES1
	                                    ,T2.ContNOIDSES2
	                                    ,T2.ContNOIDSES3
	                            FROM
	                            (
		                            SELECT  T1.FECHA
		                                    ,T1.TOTAL
		                                    ,T1.ORDEÑO
		                                    ,T1.SECAS
		                                    ,T1.HATO
		                                    ,T1.PLACT
		                                    ,T1.PPROT
		                                    ,T1.UREA
		                                    ,T1.GRASA
		                                    ,T1.CCS
		                                    ,T1.CTD
		                                    ,T1.LECPROD
		                                    ,T1.ANTIB
		                                    ,T1.X
		                                    ,T1.TOTALP
		                                    ,T1.DELORD
		                                    ,T1.VACAANTIB
		                                    ,T1.ILCA
		                                    ,T1.IC
		                                    ,T1.ILCA_P
		                                    ,T1.IC_P
		                                    ,T1.L
		                                    ,T1.MH
		                                    ,T1.PMS
		                                    ,T1.MS
		                                    ,T1.SA
		                                    ,T1.MSS
		                                    ,T1.EAS
		                                    ,T1.S
		                                    ,T1.COSTO
		                                    ,T1.PRECMS
		                                    ,CRIBA.criba1
		                                    ,CRIBA.criba2
		                                    ,CRIBA.criba3
		                                    ,CRIBA.criba4
		                                    ,T1.NOIDSES1
		                                    ,T1.NOIDSES2
		                                    ,T1.NOIDSES3
		                                    ,T1.ContNOIDSES1
		                                    ,T1.ContNOIDSES2
		                                    ,T1.ContNOIDSES3
		                            FROM
		                            (
			                            SELECT  T.FECHA
			                                    ,T.TOTAL
			                                    ,T.ORDEÑO
			                                    ,T.SECAS
			                                    ,T.HATO
			                                    ,T.PLACT
			                                    ,T.PPROT
			                                    ,T.UREA
			                                    ,T.GRASA
			                                    ,T.CCS
			                                    ,T.CTD
			                                    ,T.LECPROD
			                                    ,T.ANTIB
			                                    ,T.X
			                                    ,T.TOTALP
			                                    ,i.delord
			                                    ,T.VACAANTIB
			                                    ,T.ILCA
			                                    ,T.IC
			                                    ,T.ILCA_P
			                                    ,T.IC_P
			                                    ,IIF(T.X > 0,T.COSTO / T.X,0)      AS L
			                                    ,T.MH
			                                    ,IIF(T.MH > 0,T.MS / T.MH * 100,0) AS PMS
			                                    ,T.MS
			                                    ,T.SA
			                                    ,T.MSS
			                                    ,T.EAS
			                                    ,IIF(T.MH > 0,T.SA / T.MH * 100,0) AS S
			                                    ,T.COSTO
			                                    ,T.COSTO / T.MS                    AS PRECMS
			                                    ,T.NOIDSES1
			                                    ,T.NOIDSES2
			                                    ,T.NOIDSES3
			                                    ,T.ContNOIDSES1
			                                    ,T.ContNOIDSES2
			                                    ,T.ContNOIDSES3
			                            FROM
			                            (
				                            SELECT  m.FECHA
				                                    ,(LECFEDERAL + LECPLANTA)                                          AS TOTAL
				                                    ,VACASORDEÑA                                                       AS ORDEÑO
				                                    ,VACASSECAS                                                        AS SECAS
				                                    ,VACASHATO                                                         AS HATO
				                                    ,LACTOSA                                                           AS PLACT
				                                    ,PROTEINA                                                          AS PPROT
				                                    ,d.UREA1                                                           AS UREA
				                                    ,d.Grasa1                                                          AS GRASA
				                                    ,d.CCS1                                                            AS CCS
				                                    ,d.CTD1                                                            AS CTD
				                                    ,M.LECPROD
				                                    ,M.ANTPROD                                                         AS ANTIB
				                                    ,IIF(vacasordeña > 0,ROUND((lecprod + antprod) / vacasordeña,2),0) AS X
				                                    ,(M.LECPROD + M.ANTPROD)                                           AS TOTALP
				                                    ,M.VACAANTIB
				                                    ,M.ILCA
				                                    ,M.IC
				                                    ,M.EA
				                                    ,M.ILCA_P
				                                    ,M.IC_P
				                                    ,M.COSTO
				                                    ,M.MH
				                                    ,M.MS
				                                    ,M.SA
				                                    ,M.MSS
				                                    ,M.EAS
				                                    ,M.NOIDSES1REAL                                                    AS NOIDSES1
				                                    ,M.NOIDSES2REAL                                                    AS NOIDSES2
				                                    ,M.NOIDSES3REAL                                                    AS NOIDSES3
				                                    ,IIF(M.NOIDSES1REAL > 0,1,0)                                       AS ContNOIDSES1
				                                    ,IIF(M.NOIDSES2REAL > 0,1,0)                                       AS ContNOIDSES2
				                                    ,IIF(M.NOIDSES3REAL > 0,1,0)                                       AS ContNOIDSES3
				                            FROM MPRODUC M
				                            INNER JOIN
				                            (
					                            SELECT  FECHA
					                                    ,AVG(proteina) AS Proteina1
					                                    ,AVG(urea)     AS Urea1
					                                    ,AVG(grasa)    AS Grasa1
					                                    ,AVG(ccs)      AS CCS1
					                                    ,AVG(ctd)      AS CTD1
					                            FROM dproduc
					                            WHERE FECHA BETWEEN @inicio AND @fin
					                            GROUP BY  FECHA
				                            ) d
				                            ON m.FECHA = d.FECHA
				                            WHERE m.FECHA BETWEEN @inicio AND @fin
				                            ORDER BY m.FECHA 
			                            ) T
			                            LEFT JOIN inventario i
			                            ON i.FECHA = T.FECHA
			                            ORDER BY 1
		                            ) T1
		                            LEFT JOIN
		                            (
			                            SELECT  FECHA
			                                    ,IIF(ISNULL(AVG(nivel1)),0,AVG(nivel1)) AS criba1
			                                    ,IIF(ISNULL(AVG(nivel2)),0,AVG(nivel2)) AS criba2
			                                    ,IIF(ISNULL(AVG(nivel3)),0,AVG(nivel3)) AS criba3
			                                    ,IIF(ISNULL(AVG(nivel4)),0,AVG(nivel4)) AS criba4
			                            FROM NIVELCRIBA
			                            WHERE FECHA BETWEEN @inicio AND @fin
			                            GROUP BY  FECHA
		                            )CRIBA
		                            ON T1.FECHA = CRIBA.FECHA
	                            ) T2
	                            LEFT JOIN PRECIOSTEORICOS p
	                            ON T2.FECHA = p.FECHA
	                            ORDER BY T2.FECHA
                            ) T3";
            }
            else
            {
                query = @"SELECT  IIF(ISNULL(AVG(T3.TOTAL)),NULL,CLng(AVG(T3.TOTAL)))
                                    ,IIF(ISNULL(AVG(T3.ORDEÑO)),NULL,CINT(AVG(T3.ORDEÑO)))       AS ORDEÑO
                                    ,IIF(ISNULL(AVG(T3.SECAS)),NULL,CINT(AVG(T3.SECAS)))         AS SECAS
                                    ,IIF(ISNULL(AVG(T3.HATO)),NULL,CINT(AVG(T3.HATO)))           AS HATO
                                    ,AVG(T3.PLACT)                                               AS LACT
                                    ,AVG(T3.PPROT)                                               AS PROT
                                    ,AVG(T3.UREA)                                                AS UREA
                                    ,AVG(T3.GRASA)                                               AS GRASA
                                    ,AVG(T3.CCS)                                                 AS CCS
                                    ,IIF(ISNULL(AVG(T3.CTD)),NULL,CInt(AVG(T3.CTD)))             AS CTD
                                    ,IIF(ISNULL(AVG(T3.LECPROD)),NULL,CLng(AVG(T3.LECPROD)))     AS LECHE
                                    ,IIF(ISNULL(AVG(T3.ANTIB)),NULL,CINT(AVG(T3.ANTIB)))         AS ANTIB
                                    ,IIF(ISNULL(AVG(T3.TOTALP)),NULL,CLNG(AVG(T3.TOTALP)))       AS TOTAL2
                                    ,IIF(ISNULL(AVG(T3.DELORD)),NULL,CINT(AVG(T3.DELORD)))       AS DEL
                                    ,IIF(ISNULL(AVG(T3.VACAANTIB)),NULL,CINT(AVG(T3.VACAANTIB))) AS ANT
                                    ,AVG(T3.criba1)                                              AS N1
                                    ,AVG(T3.criba2)                                              AS N2
                                    ,AVG(T3.criba3)                                              AS N3
                                    ,AVG(T3.criba4)                                              AS N4
                                    ,IIF(ISNULL(AVG(T3.NOIDSES1)),NULL,CINT(AVG(T3.NOIDSES1)))   AS SES1
                                    ,IIF(ISNULL(AVG(T3.NOIDSES2)),NULL,CINT(AVG(T3.NOIDSES2)))   AS SES2
                                    ,IIF(ISNULL(AVG(T3.NOIDSES3)),NULL,CINT(AVG(T3.NOIDSES3)))   AS SES3
                            FROM
                            (
	                            SELECT  T2.FECHA
	                                    ,T2.TOTAL
	                                    ,T2.ORDEÑO
	                                    ,T2.SECAS
	                                    ,T2.HATO
	                                    ,T2.PLACT
	                                    ,T2.PPROT
	                                    ,T2.UREA
	                                    ,T2.GRASA
	                                    ,T2.CCS
	                                    ,T2.CTD
	                                    ,T2.LECPROD
	                                    ,T2.ANTIB
	                                    ,T2.X
	                                    ,T2.TOTALP
	                                    ,T2.DELORD
	                                    ,T2.VACAANTIB
	                                    ,CRIBA.criba1
	                                    ,CRIBA.criba2
	                                    ,CRIBA.criba3
	                                    ,CRIBA.criba4
	                                    ,T2.NOIDSES1
	                                    ,T2.NOIDSES2
	                                    ,T2.NOIDSES3
	                            FROM
	                            (
		                            SELECT  T1.FECHA
		                                    ,T1.TOTAL
		                                    ,T1.ORDEÑO
		                                    ,T1.SECAS
		                                    ,T1.HATO
		                                    ,T1.PLACT
		                                    ,T1.PPROT
		                                    ,T1.UREA
		                                    ,T1.GRASA
		                                    ,T1.CCS
		                                    ,T1.CTD
		                                    ,T1.LECPROD
		                                    ,T1.ANTIB
		                                    ,T1.X
		                                    ,T1.TOTALP
		                                    ,T1.DELORD
		                                    ,T1.VACAANTIB
		                                    ,T1.ILCA
		                                    ,T1.IC
		                                    ,T1.ILCA_P
		                                    ,T1.IC_P
		                                    ,T1.L
		                                    ,T1.MH
		                                    ,T1.PMS
		                                    ,T1.MS
		                                    ,T1.SA
		                                    ,T1.MSS
		                                    ,T1.EAS
		                                    ,T1.S
		                                    ,T1.COSTO
		                                    ,T1.PRECMS
		                                    ,CRIBA.criba1
		                                    ,CRIBA.criba2
		                                    ,CRIBA.criba3
		                                    ,CRIBA.criba4
		                                    ,T1.NOIDSES1
		                                    ,T1.NOIDSES2
		                                    ,T1.NOIDSES3
		                            FROM
		                            (
			                            SELECT  T.FECHA
			                                    ,T.TOTAL
			                                    ,T.ORDEÑO
			                                    ,T.SECAS
			                                    ,T.HATO
			                                    ,T.PLACT
			                                    ,T.PPROT
			                                    ,T.UREA
			                                    ,T.GRASA
			                                    ,T.CCS
			                                    ,T.CTD
			                                    ,T.LECPROD
			                                    ,T.ANTIB
			                                    ,T.X
			                                    ,T.TOTALP
			                                    ,i.delord
			                                    ,T.VACAANTIB
			                                    ,T.ILCA
			                                    ,T.IC
			                                    ,T.ILCA_P
			                                    ,T.IC_P
			                                    ,IIF(T.X > 0,T.COSTO / T.X,0)      AS L
			                                    ,T.MH
			                                    ,IIF(T.MH > 0,T.MS / T.MH * 100,0) AS PMS
			                                    ,T.MS
			                                    ,T.SA
			                                    ,T.MSS
			                                    ,T.EAS
			                                    ,IIF(T.MH > 0,T.SA / T.MH * 100,0) AS S
			                                    ,T.COSTO
			                                    ,T.COSTO / T.MS                    AS PRECMS
			                                    ,T.NOIDSES1
			                                    ,T.NOIDSES2
			                                    ,T.NOIDSES3
			                            FROM
			                            (
				                            SELECT  m.FECHA
				                                    ,(LECFEDERAL + LECPLANTA)                                          AS TOTAL
				                                    ,VACASORDEÑA                                                       AS ORDEÑO
				                                    ,VACASSECAS                                                        AS SECAS
				                                    ,VACASHATO                                                         AS HATO
				                                    ,LACTOSA                                                           AS PLACT
				                                    ,PROTEINA                                                          AS PPROT
				                                    ,d.UREA1                                                           AS UREA
				                                    ,d.Grasa1                                                          AS GRASA
				                                    ,d.CCS1                                                            AS CCS
				                                    ,d.CTD1                                                            AS CTD
				                                    ,M.LECPROD
				                                    ,M.ANTPROD                                                         AS ANTIB
				                                    ,IIF(vacasordeña > 0,ROUND((lecprod + antprod) / vacasordeña,2),0) AS X
				                                    ,(M.LECPROD + M.ANTPROD)                                           AS TOTALP
				                                    ,M.VACAANTIB
				                                    ,M.ILCA
				                                    ,M.IC
				                                    ,M.EA
				                                    ,M.ILCA_P
				                                    ,M.IC_P
				                                    ,M.COSTO
				                                    ,M.MH
				                                    ,M.MS
				                                    ,M.SA
				                                    ,M.MSS
				                                    ,M.EAS
				                                    ,M.NOIDSES1
				                                    ,M.NOIDSES2
				                                    ,M.NOIDSES3
				                            FROM MPRODUC M
				                            INNER JOIN
				                            (
					                            SELECT  FECHA
					                                    ,AVG(proteina) AS Proteina1
					                                    ,AVG(urea)     AS Urea1
					                                    ,AVG(grasa)    AS Grasa1
					                                    ,AVG(ccs)      AS CCS1
					                                    ,AVG(ctd)      AS CTD1
					                            FROM dproduc
					                            WHERE FECHA BETWEEN @inicio AND @fin
					                            GROUP BY  FECHA
				                            ) d
				                            ON m.FECHA = d.FECHA
				                            WHERE m.FECHA BETWEEN @inicio AND @fin
				                            ORDER BY m.FECHA 
			                            ) T
			                            LEFT JOIN inventario i
			                            ON i.FECHA = T.FECHA
			                            ORDER BY 1
		                            ) T1
		                            LEFT JOIN
		                            (
			                            SELECT  FECHA
			                                    ,IIF(ISNULL(AVG(nivel1)),0,AVG(nivel1)) AS criba1
			                                    ,IIF(ISNULL(AVG(nivel2)),0,AVG(nivel2)) AS criba2
			                                    ,IIF(ISNULL(AVG(nivel3)),0,AVG(nivel3)) AS criba3
			                                    ,IIF(ISNULL(AVG(nivel4)),0,AVG(nivel4)) AS criba4
			                            FROM NIVELCRIBA
			                            WHERE FECHA BETWEEN @inicio AND @fin
			                            GROUP BY  FECHA
		                            )CRIBA
		                            ON T1.FECHA = CRIBA.FECHA
	                            ) T2
	                            LEFT JOIN PRECIOSTEORICOS p
	                            ON T2.FECHA = p.FECHA
	                            ORDER BY T2.FECHA
                            ) T3 ";
            }

            query = query.Replace("@inicio", inicio.ToString()).Replace("@fin", fin.ToString());

            /*
            string query = "SELECT IIF(ISNULL(AVG(T3.TOTAL)),NULL,CLng(AVG(T3.TOTAL))), IIF(ISNULL(AVG(T3.ORDEÑO)),NULL,CINT(AVG(T3.ORDEÑO))) AS ORDEÑO, IIF(ISNULL(AVG(T3.SECAS)),NULL,CINT(AVG(T3.SECAS))) AS SECAS, IIF(ISNULL(AVG(T3.HATO)),NULL,CINT(AVG(T3.HATO)))  AS HATO, "
            + " AVG(T3.PLACT) AS LACT, AVG(T3.PPROT) AS PROT, AVG(T3.UREA) AS UREA, AVG(T3.GRASA) AS GRASA, "
            + " AVG(T3.CCS) AS CCS, IIF(ISNULL(AVG(T3.CTD)),NULL,CInt(AVG(T3.CTD))) AS CTD, IIF(ISNULL(AVG(T3.LECPROD)),NULL,CLng(AVG(T3.LECPROD))) AS LECHE, "
            + " IIF(ISNULL(AVG(T3.ANTIB)),NULL,CINT(AVG(T3.ANTIB))) AS ANTIB, IIF(ISNULL(AVG(T3.TOTALP)),NULL,CLNG(AVG(T3.TOTALP))) AS TOTAL2, IIF(ISNULL(AVG(T3.DELORD)),NULL,CINT(AVG(T3.DELORD))) AS DEL, IIF(ISNULL(AVG(T3.VACAANTIB)),NULL,CINT(AVG(T3.VACAANTIB))) AS ANT, "
            + " AVG(T3.criba1) AS N1, AVG(T3.criba2) AS N2, AVG(T3.criba3) AS N3, AVG(T3.criba4) AS N4, "
            + " IIF(ISNULL(AVG(T3.NOIDSES1)),NULL,CINT(AVG(T3.NOIDSES1))) AS SES1, IIF(ISNULL(AVG(T3.NOIDSES2)),NULL,CINT(AVG(T3.NOIDSES2))) AS SES2, IIF(ISNULL(AVG(T3.NOIDSES3)),NULL,CINT(AVG(T3.NOIDSES3))) AS SES3"
            + " FROM( "
            + " SELECT T2.FECHA, T2.TOTAL, T2.ORDEÑO, T2.SECAS, T2.HATO, T2.PLACT, T2.PPROT, T2.UREA, T2.GRASA, T2.CCS, T2.CTD, T2.LECPROD, T2.ANTIB, "
            + " T2.X, T2.TOTALP, T2.DELORD, T2.VACAANTIB, CRIBA.criba1, CRIBA.criba2, CRIBA.criba3, CRIBA.criba4, T2.NOIDSES1, T2.NOIDSES2, T2.NOIDSES3 "
            + " FROM( "
            + " SELECT T1.FECHA, T1.TOTAL, T1.ORDEÑO, T1.SECAS, T1.HATO, T1.PLACT, T1.PPROT, T1.UREA, T1.GRASA, T1.CCS, T1.CTD, T1.LECPROD, T1.ANTIB, "
            + " T1.X, T1.TOTALP, T1.DELORD, T1.VACAANTIB, T1.ILCA, T1.IC, T1.ILCA_P, T1.IC_P, T1.L, T1.MH, T1.PMS, T1.MS, T1.SA, T1.MSS, T1.EAS, T1.S, T1.COSTO, "
            + " T1.PRECMS, CRIBA.criba1, CRIBA.criba2, CRIBA.criba3, CRIBA.criba4, T1.NOIDSES1, T1.NOIDSES2, T1.NOIDSES3 "
            + " FROM( "
            + " SELECT T.FECHA, T.TOTAL, T.ORDEÑO, T.SECAS, T.HATO, T.PLACT, T.PPROT, T.UREA, T.GRASA, T.CCS, T.CTD, T.LECPROD, T.ANTIB, "
            + " T.X, T.TOTALP, i.delord, T.VACAANTIB, T.ILCA, T.IC, T.ILCA_P, T.IC_P, IIF(T.X > 0, T.COSTO / T.X, 0) AS L, T.MH, IIF(T.MH > 0, T.MS / T.MH * 100, 0) AS PMS, T.MS, T.SA, "
            + " T.MSS, T.EAS, IIF(T.MH > 0, T.SA / T.MH * 100, 0) AS S, T.COSTO, T.COSTO / T.MS AS PRECMS, T.NOIDSES1, T.NOIDSES2, T.NOIDSES3 "
            + " FROM( "
            + " SELECT m.FECHA, (LECFEDERAL + LECPLANTA)  AS TOTAL, VACASORDEÑA AS ORDEÑO, VACASSECAS AS SECAS, VACASHATO AS HATO, "
            + " LACTOSA AS PLACT, PROTEINA AS  PPROT, d.UREA1 AS UREA, d.Grasa1 AS GRASA, d.CCS1 AS CCS, d.CTD1 AS CTD, M.LECPROD, M.ANTPROD AS ANTIB, "
            + " IIF(vacasordeña > 0, ROUND((lecprod + antprod) / vacasordeña, 2), 0) AS X, (M.LECPROD + M.ANTPROD) AS TOTALP, "
            + " M.VACAANTIB, M.ILCA, M.IC, M.EA, M.ILCA_P, M.IC_P, M.COSTO, M.MH, M.MS, M.SA, M.MSS, M.EAS, M.NOIDSES1, M.NOIDSES2, M.NOIDSES3 "
            + " FROM MPRODUC M "
            + " INNER  JOIN( "
            + " SELECT FECHA, AVG(proteina) AS Proteina1, AVG(urea) AS Urea1, AVG(grasa) AS Grasa1, AVG(ccs) AS CCS1, AVG(ctd) AS CTD1 "
            + " FROM dproduc "
            + " WHERE FECHA BETWEEN " + inicio + " AND " + fin
            + " GROUP BY FECHA) d ON m.FECHA = d.FECHA "
            + " WHERE m.FECHA BETWEEN " + inicio + " AND " + fin
            + " ORDER BY m.FECHA) T "
            + " LEFT JOIN inventario i ON i.FECHA = T.FECHA ORDER BY 1) T1 "
            + " LEFT JOIN( "
            + " select FECHA, IIF(ISNULL(avg(nivel1)), 0, avg(nivel1)) AS criba1, IIF(ISNULL(avg(nivel2)), 0, avg(nivel2)) AS criba2, "
            + " IIF(ISNULL(avg(nivel3)), 0, avg(nivel3)) AS criba3, IIF(ISNULL(avg(nivel4)), 0, avg(nivel4))  AS criba4 "
            + " FROM NIVELCRIBA "
            + " WHERE FECHA BETWEEN " + inicio + " AND " + fin
            + " GROUP BY FECHA)CRIBA ON T1.FECHA = CRIBA.FECHA) T2 "
            + " LEFT JOIN PRECIOSTEORICOS p ON T2.FECHA = p.FECHA "
            + " ORDER BY T2.FECHA) T3 ";
            */
            conn.QueryMovsio(query, out dt);
        }

        private void Indicadores(string etapa, string campoAnimal, DateTime inicio, DateTime fin, DateTime inicio2, DateTime fin2, out DataTable dt)
        {
            DataTable dtPremezclas;
            string query = "select DISTINCT ing_descripcion FROM racion "
                       + " WHERE rac_fecha >= '" + inicio.ToString("yyyy-MM-dd HH:mm") + "' AND rac_fecha < '" + fin.ToString("yyyy-MM-dd HH:mm") + "' "
                       + " AND ran_id IN(" + ran_id.ToString() + ") AND SUBSTRING(ing_descripcion,3,2) IN('00', '01', '02') "
                       + " AND SUBSTRING(ing_descripcion,1,1) NOT IN('A','F') AND etp_id IN(" + etapa + ")";
            conn.QueryAlimento(query, out dtPremezclas);

            conn.DeleteAlimento("porcentaje_Premezcla", "");
            DataTable dtt;

            for (int i = 0; i < dtPremezclas.Rows.Count; i++)
            {
                query = "SELECT TOP(5) * FROM premezcla where pmez_racion like '" + dtPremezclas.Rows[i][0].ToString() + "'";
                conn.QueryAlimento(query, out dtt);

                if (dtt.Rows.Count == 0)
                    continue;

                CargarPremezcla(dtPremezclas.Rows[i][0].ToString(), inicio, fin);
            }

            dt = DatosInd(etapa, campoAnimal, inicio, fin, inicio2, fin2);
        }

        private DataTable DatosInd(string etapa, string campoAnimal, DateTime inicio, DateTime fin, DateTime inicio2, DateTime fin2)
        {
            DataTable dtIndicador;
            DataTable dtT, dtML;
            double racT = 0, cosT = 0, animales, media = 0, lecheF = 0, leche = 0, pms = 0, mh = 0, sob = 0;
            ColumnasIndicadores(out dtIndicador);

            string etp = "";
            switch (etapa)
            {
                case "10,11,12,13": etp = "1"; break;
                case "21": etp = "2"; break;
                case "22": etp = "4"; break;
                case "31": case "32": case "33": case "34": etp = "3"; break;
            }

            dtT = Costo(etapa, campoAnimal, inicio, fin, inicio2, fin2);
            Double.TryParse(dtT.Rows[0][0].ToString(), out racT);
            Double.TryParse(dtT.Rows[0][1].ToString(), out cosT);
            if (racT > 0)
            {
                animales = Animales(campoAnimal, inicio2, fin2);
                dtML = MediaLecheF(inicio2, fin2);
                if (dtML.Rows.Count > 0)
                {
                    Double.TryParse(dtML.Rows[0][0].ToString(), out media);
                    Double.TryParse(dtML.Rows[0][1].ToString(), out lecheF);
                }
                leche = PrecioL();
                pms = PMS(etapa, inicio, fin);
                mh = animales > 0 ? racT / animales : 0;
                sob = Sobrante(etp, inicio.AddDays(1), fin.AddDays(1));

                DataRow dr = dtIndicador.NewRow();
                dr["Animales"] = animales;
                dr["media"] = etp == "1" ? media : 0;
                dr["ilcavta"] = etp == "1" && animales > 0 && cosT > 0 ? (lecheF / animales * leche / cosT) : 0;
                dr["icventa"] = etp == "1" && animales > 0 ? (lecheF / animales * leche) - cosT : 0;
                dr["eaprod"] = etp == "1" && media > 0 && animales > 0 ? media / (pms * (racT / animales) / 100) : 0;
                dr["ilcaprod"] = etp == "1" && media > 0 && cosT > 0 ? leche * media / cosT : 0;
                dr["icprod"] = etp == "1" && media > 0 ? (leche * media) - cosT : 0;
                dr["preclprod"] = etp == "1" && media > 0 ? cosT / media : 0;
                dr["mhprod"] = animales > 0 ? racT / animales : 0;
                dr["porcmsprod"] = pms;
                dr["msprod"] = animales > 0 ? pms * (racT / animales) / 100 : 0;
                dr["saprod"] = animales > 0 ? sob / animales : 0;
                dr["mssprod"] = animales > 0 ? ((racT - sob) / animales) * pms / 100 : 0;
                dr["easprod"] = media > 0 && etp == "1" && ((racT - sob) / animales * pms / 100) > 0 ? media / ((racT - sob) / animales * pms / 100) : 0;
                dr["precprod"] = cosT;
                dr["precmsprod"] = cosT > 0 && (pms * (racT / animales) / 100) > 0 ? cosT / (pms * (racT / animales) / 100) : 0;
                dtIndicador.Rows.Add(dr);
            }
            else
            {
                DataRow dr = dtIndicador.NewRow();
                dr["Animales"] = 0;
                dr["media"] = 0;
                dr["ilcavta"] = 0;
                dr["icventa"] = 0;
                dr["ilcaprod"] = 0;
                dr["icprod"] = 0;
                dr["preclprod"] = 0;
                dr["mhprod"] = 0;
                dr["porcmsprod"] = 0;
                dr["msprod"] = 0;
                dr["saprod"] = 0;
                dr["mssprod"] = 0;
                dr["easprod"] = 0;
                dr["precprod"] = 0;
                dr["precmsprod"] = 0;
                dtIndicador.Rows.Add(dr);
            }

            return dtIndicador;
        }
        private double Sobrante(string etp, DateTime inicio, DateTime fin)
        {
            double sob = 0;
            DataTable dt;
            string query = "SELECT ISNULL(SUM(rac_mh)/ DATEDIFF(DAY, '" + inicio.ToString("yyyy-MM-dd HH:mm") + "', '" + fin.ToString("yyyy-MM-dd HH:mm") + "'),0) AS Peso "
                            + " FROM racion where rac_fecha >= '" + inicio.ToString("yyyy-MM-dd HH:mm") + "' AND rac_fecha < '" + fin.ToString("yyyy-MM-dd HH:mm") + "' AND rac_descripcion like '%SOB%' "
                            + " AND ing_descripcion like '" + etp + "%' AND ran_id IN(" + ran_id + ") AND SUBSTRING(ing_descripcion,3,2) NOT IN('00','01','02','90')";
            conn.QueryAlimento(query, out dt);

            if (dt.Rows.Count > 0)
                Double.TryParse(dt.Rows[0][0].ToString(), out sob);
            return sob;
        }
        private DataTable MediaLecheF(DateTime inicio, DateTime fin)
        {
            DataTable dt;
            string query = "SELECT IIF(SUM(ia.ia_vacas_ord)>0,ISNULL((SUM(m.med_produc)/SUM(ia.ia_vacas_ord)),0),0) , ISNULL(SUM(m.med_lecfederal + m.med_lecplanta) / COUNT(DISTINCT med_fecha),0) "
                                + " FROM media m LEFT JOIN inventario_afi ia ON ia.ran_id = m.ran_id AND ia.ia_fecha = m.med_fecha "
                                + " where m.ran_id IN(" + ran_id + ") AND med_fecha BETWEEN '" + inicio.ToString("yyyy-MM-dd") + "' AND '" + fin.ToString("yyyy-MM-dd") + "' ";
            conn.QueryAlimento(query, out dt);
            return dt;
        }
        private double PrecioL()
        {
            double precio = 0;
            DataTable dt;
            string query = "SELECT TOP(1) hl_precio FROM historico_leche WHERE ran_id = " + ran_id + " ORDER BY hl_fecha_reg desc";
            conn.QueryAlimento(query, out dt);
            Double.TryParse(dt.Rows[0][0].ToString(), out precio);
            return precio;
        }
        private void ColumnasIndicadores(out DataTable dt)
        {
            dt = new DataTable();
            dt.Columns.Add("Animales").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("media").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("ilcavta").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("icventa").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("eaprod").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("ilcaprod").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("icprod").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("preclprod").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("mhprod").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("porcmsprod").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("msprod").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("saprod").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("mssprod").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("easprod").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("precprod").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("precmsprod").DataType = System.Type.GetType("System.Double");
        }
        private DataTable Costo(string etapa, string campo, DateTime inicio, DateTime fin, DateTime inicio2, DateTime fin2)
        {
            double v = 0;
            DataTable dt;
            ColumnasDT(out dt);
            int vacas = Animales(campo, inicio2, fin2);
            string sob = Sobrantes();
            DataTable dt1;
            string query = "SELECT x.Ing, SUM(X.PesoH *X.Precio)/SUM(X.PesoH) AS Precio, SUM(X.PesoS) / SUM(X.PesoH)*100 AS PMS, SUM(X.PesoH) AS PesoH, SUM(X.PesoS) AS PesoS "
            + " FROM( "
            + " SELECT R.Clave, R.Ing, ISNULL(i.ing_precio_sie, 0) AS Precio, ISNULL(it.ingt_porcentaje_ms, 0) AS PMS, "
            + " SUM(R.Peso) / DATEDIFF(DAY, '" + inicio.ToString("yyyy-MM-dd HH:mm") + "', '" + fin.ToString("yyyy-MM-dd HH:mm") + "')  AS PesoH, "
            + " (SUM(R.Peso) * ISNULL(it.ingt_porcentaje_ms, 0) / 100) / DATEDIFF(DAY, '" + inicio.ToString("yyyy-MM-dd HH:mm") + "', '" + fin.ToString("yyyy-MM-dd HH:mm") + "')  AS PesoS "
            + " FROM( "
            + " SELECT ran_id AS Ran, r.ing_clave AS Clave, r.ing_descripcion AS Ing, SUM(rac_mh) AS Peso "
            + " FROM racion r "
            + " WHERE ran_id IN(" + ran_id + ")  AND rac_fecha >= '" + inicio.ToString("yyyy-MM-dd HH:mm") + "' AND rac_fecha < '" + fin.ToString("yyyy-MM-dd HH:mm") + "' AND etp_id IN(" + etapa + ") "
            + " AND SUBSTRING(ing_clave, 1, 4) IN('ALAS', 'ALFO') GROUP BY ran_id, ing_clave, ing_descripcion "
            + " UNION "
            + " SELECT T.Ran, T.Clave, T.Ing, SUM(T.Peso) "
            + " FROM( "
            + " SELECT T1.Ran, IIF(T2.Pmez IS NULL, T1.Clave, T2.Clave) AS Clave, IIF(T2.Pmez IS NULL, T1.Ing, T2.Ing) AS Ing, IIF(T2.Pmez IS NULL, T1.Peso, T1.Peso * T2.Porc) AS Peso "
            + " FROM( "
            + " SELECT T1.Ran, T2.Clave, T2.Ing, (T1.Peso * T2.Porc) AS Peso "
            + " FROM( "
            + " SELECT ran_id As Ran, ing_descripcion AS Pmz, SUM(rac_mh) AS Peso "
            + " FROM racion "
            + " WHERE ran_id IN(" + ran_id + ")  AND rac_fecha >= '" + inicio.ToString("yyyy-MM-dd HH:mm") + "' AND rac_fecha < '" + fin.ToString("yyyy-MM-dd HH:mm") + "'  AND etp_id IN(" + etapa + ") "
            + " AND ISNUMERIC(SUBSTRING(ing_descripcion, 1, 1)) > 0 AND SUBSTRING(ing_descripcion, 3, 2) IN('00', '01', '02') "
            + " GROUP BY ran_id, ing_descripcion) T1 "
            + " LEFT JOIN(SELECT pmez_descripcion AS Pmez, ing_clave AS Clave, ing_descripcion AS Ing, pmez_porcentaje AS Porc FROM porcentaje_Premezcla)T2 ON T1.Pmz = T2.Pmez) T1 "
            + " LEFT JOIN(SELECT pmez_descripcion AS Pmez, ing_clave AS Clave, ing_descripcion AS Ing, pmez_porcentaje AS Porc FROM porcentaje_Premezcla)T2 ON T1.Ing = T2.Pmez) T "
            + " GROUP BY T.Ran, T.Clave, T.Ing "
            + " UNION "
            + " SELECT ran_id, ing_clave, ing_descripcion, SUM(rac_mh) "
            + " FROM racion "
            + " WHERE rac_fecha >= '" + inicio.ToString("yyyy-MM-dd HH:mm") + "' AND rac_fecha < '" + fin.ToString("yyyy-MM-dd HH:mm") + "' AND ran_id IN(" + ran_id + ")  AND SUBSTRING(ing_descripcion, 1, 1) NOT IN('A', 'F', 'W') "
            + " AND SUBSTRING(ing_descripcion, 3, 2) NOT IN('00', '01', '02')  AND etp_id IN(" + etapa + ") "
            + " GROUP BY ran_id, ing_clave, ing_descripcion "
            + " UNION "
            + " SELECT ran_id, ing_clave, ing_descripcion, SUM(rac_mh) "
            + " FROM racion "
            + " WHERE rac_fecha >= '" + inicio.ToString("yyyy-MM-dd HH:mm") + "' AND rac_fecha < '" + fin.ToString("yyyy-MM-dd HH:mm") + "' AND ran_id IN(" + ran_id + ")  AND ing_descripcion IN('Agua', 'Water') "
            + " AND etp_id IN(" + etapa + ")  GROUP BY ran_id, ing_clave, ing_descripcion) R "
            + " LEFT JOIN ingrediente i ON i.ing_clave = R.Clave AND i.ing_descripcion = R.Ing AND i.ran_id = R.Ran "
            + " LEFT JOIN ingrediente_tracker it ON it.ingt_clave = R.Clave AND R.Ing = it.ingt_descripcion AND R.Ran = it.ran_id "
            + " GROUP BY R.Ran, R.Clave, R.Ing, i.ing_precio_sie, it.ingt_porcentaje_ms) X "
            + " WHERE X.PesoH > 0 GROUP BY x.Ing";
            conn.QueryAlimento(query, out dt1);

            DataTable dtTemp; ColumnasDT(out dtTemp);
            double xvaca, s_xvaca, totalR = 0, costoT = 0, txvaca = 0, tsxvaca = 0, costo;
            double mh, ms, pms, precio;

            for (int i = 0; i < dt1.Rows.Count; i++)
            {
                precio = Convert.ToDouble(dt1.Rows[i][1]);
                pms = Convert.ToDouble(dt1.Rows[i][2]);
                mh = Convert.ToDouble(dt1.Rows[i][3]); totalR += mh;
                ms = Convert.ToDouble(dt1.Rows[i][4]);
                xvaca = vacas > 0 ? mh / vacas : 0;
                s_xvaca = vacas > 0 ? ms / vacas : 0;
                txvaca += xvaca;
                tsxvaca += s_xvaca;
                costoT += precio * xvaca;
                DataRow dr = dtTemp.NewRow();
                dr["ingrediente"] = dt1.Rows[i][0].ToString();
                dr["precioIng"] = dt1.Rows[i][1].ToString();
                dr["xvaca"] = xvaca;
                dr["TOTAL"] = mh;
                dr["COSTO"] = precio * xvaca;
                dr["PRECIO"] = precio * mh;
                dr["s_precioIng"] = pms > 0 ? precio * 100 / pms : 0;
                dr["s_xvaca"] = s_xvaca;
                dr["s_TOTAL"] = ms;
                dr["s_COSTO"] = (pms > 0 ? precio * 100 / pms : 0) * s_xvaca;
                dr["s_PRECIO"] = (pms > 0 ? precio * 100 / pms : 0) * s_xvaca;
                dr["PMS"] = pms;
                dtTemp.Rows.Add(dr);
            }

            for (int i = 0; i < dtTemp.Rows.Count; i++)
            {
                xvaca = Convert.ToDouble(dtTemp.Rows[i]["xvaca"]);
                s_xvaca = Convert.ToDouble(dtTemp.Rows[i]["s_xvaca"]);
                mh = Convert.ToDouble(dtTemp.Rows[i]["TOTAL"]);
                costo = Convert.ToDouble(dtTemp.Rows[i]["COSTO"]);
                dtTemp.Rows[i]["porcvaca"] = xvaca / txvaca * 100;
                dtTemp.Rows[i]["porccosto"] = costo / costoT * 100;
                dtTemp.Rows[i]["s_porcvaca"] = s_xvaca / tsxvaca * 100;
                dtTemp.Rows[i]["s_porccosto"] = costo / costoT * 100;
            }

            string ing, ingA;

            for (int i = 0; i < dtTemp.Rows.Count; i++)
            {
                v += Convert.ToDouble(dtTemp.Rows[i]["COSTO"]);
            }
            DataTable dtC = new DataTable();
            dtC.Columns.Add("racion");
            dtC.Columns.Add("costo");

            DataRow drc = dtC.NewRow();
            drc["racion"] = totalR > 0 ? totalR : 0;
            drc["costo"] = v > 0 ? v : 0;
            dtC.Rows.Add(drc);

            return dtC;

        }
        private void ColumnasDT(out DataTable dt)
        {
            dt = new DataTable();
            dt.Columns.Add("ingrediente").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("precioIng").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("xvaca").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("porcvaca").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("TOTAL").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("COSTO").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("porccosto").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIO").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("s_precioIng").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("s_xvaca").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("s_porcvaca").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("s_TOTAL").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("s_COSTO").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("s_porccosto").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("s_PRECIO").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PMS").DataType = System.Type.GetType("System.Double");
        }
        private int Animales(string campo, DateTime inicio, DateTime fin)
        {
            DataTable dt;
            int animales = 0;
            string query = "SELECT ROUND(SUM(CONVERT(FLOAT," + campo + ")) / COUNT(DISTINCT ia_fecha), 0 ) AS Vacas FROM inventario_afi WHERE ran_id IN( " + ran_id + ") AND ia_fecha BETWEEN '" + inicio.ToString("yyyy-MM-dd") + "' AND '" + fin.ToString("yyyy-MM-dd") + "'";
            conn.QueryAlimento(query, out dt);
            if (dt.Rows.Count > 0)
                Int32.TryParse(dt.Rows[0][0].ToString(), out animales);

            return animales;
        }
        private string Sobrantes()
        {
            DataTable dt;
            string sobrantes = "";
            string query = "SELECT description FROM ds_ingredient WHERE is_active = 1 AND is_deleted = 0 AND substring(description from 1 for 1) not in ('A','F','W') "
                    + "  AND SUBSTRING(description from 3 for 2) not in('00','01','02','90') ";
            conn.QueryTracker(query, out dt);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                sobrantes += "'" + dt.Rows[i][0].ToString() + "',";
            }

            sobrantes = sobrantes.Length > 0 ? sobrantes.Substring(0, sobrantes.Length - 1) : "''";
            return sobrantes;
        }
        private double PMS(string etapa, DateTime inicio, DateTime fin)
        {
            double v = 0;
            DataTable dt;
            string sob = Sobrantes();
            string query = @"SELECT  SUM(R.PesoS) / SUM(R.PesoH) * 100
                            FROM
                            (
                                SELECT  x.Ing
                                       , SUM(X.PesoH * X.Precio) / SUM(X.PesoH)         AS Precio
                                       , ISNULL(SUM(X.PesoS) / SUM(X.PesoH) * 100, 0) AS PMS
                                       , SUM(X.PesoH)                                AS PesoH
                                       , ISNULL(SUM(X.PesoS), 0)                      AS PesoS
                                FROM
                                (
                                    SELECT  R.Clave
                                           , R.Ing
                                           , ISNULL(i.ing_precio_sie, 0)                                         AS Precio
                                           , ISNULL(SUM(R.PesoS) / SUM(R.PesoH), 0)                              AS PMS
                                           , SUM(R.PesoH) / DATEDIFF(DAY, @FechaInicio, @FFinal) AS PesoH
                                           , SUM(R.PesoS) / DATEDIFF(DAY, @FechaInicio, @FFinal) AS PesoS
                                    FROM
                                    (
                                        SELECT  ran_id            AS Ran
                                               , r.ing_clave       AS Clave
                                               , r.ing_descripcion AS Ing
                                               , SUM(rac_mh)       AS PesoH
                                               , SUM(rac_ms)       AS PesoS
                                        FROM racion r
                                        WHERE ran_id IN(@RanId)
                                        AND rac_fecha >= @FechaInicio
                                        AND rac_fecha < @FFinal
                                        AND etp_id IN (@Etapa)
                                        AND SUBSTRING(ing_clave, 1, 4) IN('ALAS', 'ALFO')
                                        GROUP BY  ran_id
                                                 , ing_clave
                                                 , ing_descripcion UNION
                                        SELECT  T.Ran
                                               , T.Clave
                                               , T.Ing
                                               , SUM(T.Peso)
                                               , SUM(T.PesoS)
                                        FROM
                                        (
                                            SELECT  T1.Ran
                                                   , IIF(T2.Pmez IS NULL, T1.Clave, T2.Clave)               AS Clave
                                                   , IIF(T2.Pmez IS NULL, T1.Ing, T2.Ing)                   AS Ing
                                                   , IIF(T2.Pmez IS NULL, T1.Peso, T1.Peso * T2.Porc)       AS Peso
                                                   , IIF(T2.Pmez IS NULL, T1.PesoS, T1.Peso * T2.PorcSeca) AS PesoS
                                            FROM
                                            (
                                                SELECT  T1.Ran
                                                       , T2.Clave
                                                       , T2.Ing
                                                       , (T1.Peso * T2.Porc)      AS Peso
                                                       , (T1.PesoS * T2.PorcSeca) AS PesoS
                                                FROM
                                                (
                                                    SELECT  ran_id          AS Ran
                                                           , ing_descripcion AS Pmz
                                                           , SUM(rac_mh)     AS Peso
                                                           , SUM(rac_ms)     AS PesoS
                                                    FROM racion
                                                    WHERE ran_id IN(@RanId)
                                                    AND rac_fecha >= @FechaInicio
                                                    AND rac_fecha < @FFinal
                                                    AND etp_id IN (@Etapa)
                                                    AND ISNUMERIC(SUBSTRING(ing_descripcion, 1, 1)) > 0
                                                    AND SUBSTRING(ing_descripcion, 3, 2) IN('00', '01', '02')
                                                    GROUP BY  ran_id
                                                             , ing_descripcion
                                                ) T1
                                                LEFT JOIN
                                                (
                                                    SELECT  pmez_descripcion     AS Pmez
                                                           , ing_clave            AS Clave
                                                           , ing_descripcion      AS Ing
                                                           , pmez_porcentaje      AS Porc
                                                           , pmez_porcentaje_seca AS PorcSeca
                                                    FROM porcentaje_Premezcla
                                                )T2
                                                ON T1.Pmz = T2.Pmez
                                            ) T1
                                            LEFT JOIN
                                            (
                                                SELECT  pmez_descripcion     AS Pmez
                                                       , ing_clave            AS Clave
                                                       , ing_descripcion      AS Ing
                                                       , pmez_porcentaje      AS Porc
                                                       , pmez_porcentaje_seca AS PorcSeca
                                                FROM porcentaje_Premezcla
                                            )T2
                                            ON T1.Ing = T2.Pmez
                                        ) T
                                        GROUP BY  T.Ran
                                                 , T.Clave
                                                 , T.Ing UNION
                                        SELECT  ran_id
                                               , ing_clave
                                               , ing_descripcion
                                               , SUM(rac_mh)
                                               , SUM(rac_ms)
                                        FROM racion
                                        WHERE rac_fecha >= @FechaInicio
                                        AND rac_fecha < @FFinal
                                        AND ran_id IN(@RanId)
                                        AND SUBSTRING(ing_descripcion, 1, 1) NOT IN('A', 'F', 'W')
                                        AND SUBSTRING(ing_descripcion, 3, 2) NOT IN('00', '01', '02')
                                        AND etp_id IN (@Etapa)
                                        GROUP BY  ran_id
                                                 , ing_clave
                                                 , ing_descripcion UNION
                                        SELECT  ran_id
                                               , ing_clave
                                               , ing_descripcion
                                               , SUM(rac_mh)
                                               , SUM(rac_ms)
                                        FROM racion
                                        WHERE rac_fecha >= @FechaInicio
                                        AND rac_fecha < @FFinal
                                        AND ran_id IN(@RanId)
                                        AND ing_descripcion IN('Agua', 'Water')
                                        AND etp_id IN (@Etapa)
                                        GROUP BY  ran_id
                                                 , ing_clave
                                                 , ing_descripcion
                                    ) R
                                    LEFT JOIN ingrediente i
                                    ON i.ing_clave = R.Clave AND i.ing_descripcion = R.Ing AND i.ran_id = R.Ran
                                    LEFT JOIN ingrediente_tracker it
                                    ON it.ingt_clave = R.Clave AND R.Ing = it.ingt_descripcion AND R.Ran = it.ran_id
                                    GROUP BY  R.Ran
                                             , R.Clave
                                             , R.Ing
                                             , i.ing_precio_sie
                                             , it.ingt_porcentaje_ms
                                ) X
                                WHERE X.PesoH > 0
                                GROUP BY  x.Ing
                            ) R";
            query = query.Replace("@FFinal", "'" + fin.ToString("yyyy-MM-dd HH:mm") + "'")
                .Replace("@FechaInicio", "'" + inicio.ToString("yyyy-MM-dd HH:mm") + "'")
                .Replace("@RanId", ran_id.ToString())
                .Replace("@Etapa", etapa);
            conn.QueryAlimento(query, out dt);
            if (dt.Rows[0][0] != DBNull.Value)
            {
                v = dt.Rows.Count > 0 ? Convert.ToDouble(dt.Rows[0][0]) : 0;
            }

            return v;
        }

        private void CargarPremezcla(string premezcla, DateTime inicio, DateTime fin)
        {
            try
            {
                DateTime fRacion, fIng;
                DateTime fin2 = inicio.AddDays(1);
                DateTime fpmI = inicio, fpmF = new DateTime();
                int temp = 0;
                DataTable dt, dtAux;
                DataTable dt1 = new DataTable();
                string pmz, clave, ingrediente, valores = "", prmz, query;
                double porcentaje, porcentajesecas;
                int repeticiones = 0;
                prmz = premezcla[2].ToString() + premezcla[3];
                int diferencia = 0;
                TimeSpan ts = (inicio - Convert.ToDateTime("01 /01/1900"));
                int juliandayIni = ts.Days + 2;
                TimeSpan ts2 = (fin - Convert.ToDateTime("01/01/1900"));
                int juliandayFin = ts2.Days + 2;
                diferencia = juliandayFin - juliandayIni;
                query = "SELECT * FROM porcentaje_Premezcla WHERE pmez_descripcion like '" + premezcla + "'";
                conn.QueryAlimento(query, out dtAux);

                if (dtAux.Rows.Count == 0)
                {
                    if (prmz == "01")
                    {
                        query = @"SELECT
	                                   T1.Pmz
                                      ,T1.Clave
                                      ,T1.Ing
	                                  ,IIF(T2.Total > 0,  T1.Peso / T2.Total, 0)
	                                  ,IIF(SEC.Peso > 0, SEC2.Peso / SEC.Peso , 0)
                                FROM 
                                (
	                                SELECT  T.pmez_racion    AS Pmz
	                                       ,T.ing_clave      AS Clave
	                                       ,T.ing_nombre     AS Ing
	                                       ,SUM(T.pmez_peso) AS Peso
	                                FROM 
	                                (
		                                SELECT  DISTINCT *
		                                FROM premezcla
		                                WHERE pmez_racion  LIKE '" + premezcla + @"'
                                        AND pmez_fecha >= '" + inicio.ToString("yyyy-MM-dd HH:mm") + @"'
                                        AND pmez_fecha < '" + fin.ToString("yyyy-MM-dd HH:mm") + @"'
	                                ) T
	                                GROUP BY  pmez_racion
	                                         ,ing_clave
	                                         ,ing_nombre 
                                ) T1
                                LEFT JOIN 
                                (
	                                SELECT  T.pmez_racion    AS Pmz
	                                       ,SUM(T.pmez_peso) AS Total
	                                FROM 
	                                (
		                                SELECT  DISTINCT *
		                                FROM premezcla
		                                WHERE pmez_racion  LIKE '" + premezcla + @"'
                                        AND pmez_fecha >= '" + inicio.ToString("yyyy-MM-dd HH:mm") + @"'
                                        AND pmez_fecha < '" + fin.ToString("yyyy-MM-dd HH:mm") + @"'
	                                ) T
		                                GROUP BY  T.pmez_racion
                                )       T2 ON T1.Pmz = T2.Pmz
                                LEFT JOIN( 
		                                SELECT rac_descripcion AS Pmez, SUM(rac_ms) AS Peso 
		                                FROM racion 
		                                WHERE rac_fecha >= '" + inicio.ToString("yyyy-MM-dd HH:mm") + @"'
		                                AND rac_fecha  < '" + fin.ToString("yyyy-MM-dd HH:mm") + @"'
		                                AND rac_descripcion LIKE '" + premezcla + @"' 
		                                GROUP BY rac_descripcion
                                )
		                                SEC ON SEC.Pmez = T1.Pmz
                                LEFT JOIN( 
		                                SELECT rac_descripcion AS Pmez, ing_clave AS Clave, SUM(rac_ms) AS Peso 
			                                FROM racion 
			                                WHERE rac_fecha >= '" + inicio.ToString("yyyy-MM-dd HH:mm") + @"'
		                                    AND rac_fecha  < '" + fin.ToString("yyyy-MM-dd HH:mm") + @"'
		                                    AND rac_descripcion LIKE '" + premezcla + @"' 
		                                GROUP BY rac_descripcion, ing_clave
                                )
                                SEC2 ON SEC2.Clave = T1.Clave";
                        conn.QueryAlimento(query, out dt);

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            pmz = dt.Rows[i][0].ToString();
                            clave = dt.Rows[i][1].ToString();
                            ingrediente = dt.Rows[i][2].ToString();
                            porcentaje = Convert.ToDouble(dt.Rows[i][3]);
                            porcentajesecas = Convert.ToDouble(dt.Rows[i][4]);
                            valores += "('" + pmz + "','" + clave + "','" + ingrediente + "'," + porcentaje + "," + porcentajesecas + "),";
                        }
                        if (valores.Length > 0)
                        {
                            valores = valores.Substring(0, valores.Length - 1);
                            conn.InsertMasivAlimento("porcentaje_Premezcla", valores);
                        }
                    }
                    else
                    {
                        query = "SELECT T1.Premezcla, T1. Fecha AS PMIng, ISNULL(T2. Fecha, T3.Fecha) AS PMRac "
                                   + " FROM( "
                                   + " SELECT ing_descripcion AS Premezcla, MIN(rac_fecha) AS Fecha "
                                   + " FROM racion "
                                   + " WHERE rac_fecha >= '" + inicio.ToString("yyyy-MM-dd HH:mm") + "' AND rac_fecha < '" + fin.ToString("yyyy-MM-dd HH:mm") + "' "
                                   + " AND ing_descripcion like '" + premezcla + "'"
                                   + " GROUP BY ing_descripcion)T1 "
                                   + " LEFT JOIN( "
                                   + " SELECT rac_descripcion AS Premezcla, MIN(rac_fecha)  AS Fecha "
                                   + " FROM racion "
                                   + " WHERE rac_fecha >= '" + inicio.ToString("yyyy-MM-dd HH:mm") + "' AND rac_fecha < '" + fin.ToString("yyyy-MM-dd HH:mm") + "' "
                                   + " AND rac_descripcion like '" + premezcla + "' "
                                   + " GROUP BY rac_descripcion) T2 ON T1.Premezcla = T2.Premezcla "
                                   + " LEFT JOIN( "
                                   + " SELECT rac_descripcion AS Premezcla, MAX(rac_fecha)  AS Fecha "
                                   + " FROM racion "
                                   + " WHERE rac_fecha < '" + inicio.ToString("yyyy-MM-dd HH:mm") + "' AND rac_descripcion like '" + premezcla + "' "
                                   + " GROUP BY rac_descripcion "
                                   + " )T3 ON T1.Premezcla = T3.Premezcla";
                        conn.QueryAlimento(query, out dt);

                        fIng = Convert.ToDateTime(dt.Rows[0][1]);
                        fRacion = dt.Rows[0][2] == DBNull.Value ? fin : Convert.ToDateTime(dt.Rows[0][2]);
                        int comparacion = DateTime.Compare(fRacion, fIng);

                        if (comparacion == 1)
                        {
                            do
                            {
                                if (repeticiones == 30)
                                    break;

                                fpmI = inicio.AddDays(temp);
                                fpmF = fin;

                                //fpmI = inicio.AddDays(-1); fpmI = fpmI.AddDays(temp);
                                //fpmF = fin2.AddDays(-1); fpmF = fpmF.AddDays(temp);

                                query = " SELECT * FROM premezcla WHERE pmez_racion like '" + premezcla + "' "
                                    + " AND pmez_fecha >= '" + fpmI.ToString("yyyy-MM-dd HH:mm") + "' AND pmez_fecha < '" + fpmF.ToString("yyyy-MM-dd HH:mm") + "' ";
                                conn.QueryAlimento(query, out dt1);
                                temp--;

                                repeticiones++;
                            }
                            while (dt1.Rows.Count == 0);
                        }
                        else
                        {
                            if (fRacion.Hour < inicio.Hour)
                            {
                                fpmI = new DateTime(fRacion.Year, fRacion.Month, fRacion.Day, inicio.Hour, 0, 0);
                                fpmI = fpmI.AddDays(-1);
                            }
                            else
                                fpmI = new DateTime(fRacion.Year, fRacion.Month, fRacion.Day, inicio.Hour, 0, 0);
                        }

                        DataTable dtsPM;
                        query = "SELECT DISTINCT ing_nombre "
                                + "FROM premezcla "
                                + " WHERE pmez_racion like '" + premezcla + "'"
                                + " AND pmez_fecha >= '" + fpmI.ToString("yyyy-MM-dd HH:mm") + "' AND pmez_fecha < '" + fin.ToString("yyyy-MM-dd HH:mm") + "' "
                                + " AND ISNUMERIC(SUBSTRING(ing_nombre,1,1)) > 0 "
                                + " AND SUBSTRING(ing_nombre,3,2) IN('00', '01', '02')";
                        conn.QueryAlimento(query, out dtsPM);

                        DataTable dtV;
                        //DiasPremezcla(premezcla, fpmI, fin);
                        for (int i = 0; i < dtsPM.Rows.Count; i++)
                        {
                            query = "SELECT * FROM premezcla WHERE pmez_racion like '" + dtsPM.Rows[i][0].ToString() + "'";
                            conn.QueryAlimento(query, out dtV);

                            if (dtV.Rows.Count == 0)
                                continue;

                            SupraMezcla(dtsPM.Rows[i][0].ToString(), fpmI, fin);
                        }

                        if (diferencia == 1)
                        {
                            query = @"INSERT INTO porcentaje_Premezcla 
                                           SELECT  T1.PMEZ
                                           ,T1.Clave
                                           ,T1.Ing
                                           ,IIF(T2.TOTALH > 0, T1.PESOH / T2.TOTALH, 0)
										   ,IIF(T1.PESOH > 0, PESOS*100 /PESOH, 0) * IIF(T2.TOTALH > 0, (T1.PESOH / T2.TOTALH)/100, 0) AS PORC_MATERIASECA
                                    FROM
                                    (
                                        SELECT  pmez_racion                               AS PMEZ
                                               , ing_clave                                 AS Clave
                                               , ing_nombre                                AS Ing
                                               , SUM(pmez_peso)                            AS PESOH
                                               , SUM(pmez_peso * pmez_porcentaje_ms / 100) AS PESOS
                                        FROM premezcla
                                        WHERE pmez_fecha >= @fechaIni
                                        AND pmez_fecha < @FechaFin AND pmez_racion = @premezcla
                                        GROUP BY  pmez_racion
                                                 , ing_clave
                                                 , ing_nombre
                                    ) T1
                                    LEFT JOIN
                                    (
                                        SELECT  pmez_racion                               AS PMEZ
                                               , SUM(pmez_peso)                            AS TOTALH
                                               , SUM(pmez_peso * pmez_porcentaje_ms / 100) AS TOTALS
                                        FROM premezcla
                                        WHERE pmez_fecha >= @fechaIni
                                        AND pmez_fecha < @FechaFin
                                        AND pmez_racion = @premezcla
                                        GROUP BY  pmez_racion
                                    ) T2
                                    ON T1.PMEZ = T2.PMEZ";
                            query = query.Replace("@fechaIni", "'" + fpmI.ToString("yyyy-MM-dd HH:mm") + "'").Replace("@FechaFin", "'" + fin.ToString("yyyy-MM-dd HH:mm") + "'").Replace("@premezcla", "'" + premezcla + "'");
                            //   query = "INSERT INTO porcentaje_Premezcla "
                            //+ " SELECT T1.Pmez, T1.Clave, T1.Ing, IIF(T2.Peso > 0, T1.Peso / T2.Peso,0), IIF( T2.PesoSeco > 0, T1.PesoSeco / T2.PesoSeco , 0) "
                            //+ " FROM( "
                            //+ " SELECT rac_descripcion AS Pmez, ing_clave AS Clave, ing_descripcion AS Ing, SUM(rac_mh) AS Peso , SUM(rac_ms) AS PesoSeco "
                            //+ " FROM racion "
                            //+ " WHERE rac_fecha >= '" + fpmI.ToString("yyyy-MM-dd HH:mm") + "' AND rac_fecha < '" + fin.ToString("yyyy-MM-dd HH:mm") + "' "
                            //+ " AND rac_descripcion LIKE '" + premezcla + "' "
                            //+ " GROUP BY rac_descripcion, ing_clave, ing_descripcion)T1 "
                            //+ " LEFT JOIN( "
                            //+ " SELECT rac_descripcion AS Pmez, SUM(rac_mh) AS Peso , SUM(rac_ms) AS PesoSeco"
                            //+ " FROM racion "
                            //+ " WHERE rac_fecha >= '" + fpmI.ToString("yyyy-MM-dd HH:mm") + "' AND rac_fecha < '" + fin.ToString("yyyy-MM-dd HH:mm") + "' "
                            //+ " AND rac_descripcion LIKE '" + premezcla + "' "
                            //+ " GROUP BY rac_descripcion"
                            //+ " ) T2 ON T1.Pmez = T2.Pmez";
                            conn.InsertSelecttAlimento(query);
                        }
                        else
                        {
                            int temp2 = 0, repeticiones2 = 0;
                            fpmI = inicio;
                            do
                            {
                                if (repeticiones2 == 30)
                                    break;

                                //fpmI = inicio.AddDays(-1); fpmI = fpmI.AddDays(temp);
                                //fpmF = fin2.AddDays(-1); fpmF = fpmF.AddDays(temp);

                                query = " SELECT * FROM premezcla WHERE pmez_racion like '" + premezcla + "' "
                                    + " AND pmez_fecha >= '" + fpmI.ToString("yyyy-MM-dd HH:mm") + "' AND pmez_fecha < '" + fin.ToString("yyyy-MM-dd HH:mm") + "' ";
                                conn.QueryAlimento(query, out dt1);
                                fpmI = fpmI.AddDays(temp2);//En caso de emergencia remplazar temp2 con -1
                                fpmF = fin;
                                temp2--;

                                repeticiones2++;
                            }
                            while (dt1.Rows.Count == 0);
                            query = @"INSERT INTO porcentaje_Premezcla 
                                           SELECT  T1.PMEZ
                                           ,T1.Clave
                                           ,T1.Ing
                                           ,IIF(T2.TOTALH > 0, T1.PESOH / T2.TOTALH, 0)
                                           ,IIF(T1.PESOH > 0, PESOS*100 /PESOH, 0) * IIF(T2.TOTALH > 0, (T1.PESOH / T2.TOTALH)/100, 0) AS PORC_MATERIASECA
                                    FROM
                                    (
                                        SELECT  pmez_racion                               AS PMEZ
                                               , ing_clave                                 AS Clave
                                               , ing_nombre                                AS Ing
                                               , SUM(pmez_peso)                            AS PESOH
                                               , SUM(pmez_peso * pmez_porcentaje_ms / 100) AS PESOS
                                        FROM premezcla
                                        WHERE pmez_fecha >= @fechaIni
                                        AND pmez_fecha < @FechaFin AND pmez_racion = @premezcla
                                        GROUP BY  pmez_racion
                                                 , ing_clave
                                                 , ing_nombre
                                    ) T1
                                    LEFT JOIN
                                    (
                                        SELECT  pmez_racion                               AS PMEZ
                                               , SUM(pmez_peso)                            AS TOTALH
                                               , SUM(pmez_peso * pmez_porcentaje_ms / 100) AS TOTALS
                                        FROM premezcla
                                        WHERE pmez_fecha >= @fechaIni
                                        AND pmez_fecha < @FechaFin
                                        AND pmez_racion = @premezcla
                                        GROUP BY  pmez_racion
                                    ) T2
                                    ON T1.PMEZ = T2.PMEZ";
                            query = query.Replace("@fechaIni", "'" + fpmI.ToString("yyyy-MM-dd HH:mm") + "'").Replace("@FechaFin", "'" + fin.ToString("yyyy-MM-dd HH:mm") + "'").Replace("@premezcla", "'" + premezcla + "'");
                            //query = "INSERT INTO porcentaje_Premezcla "
                            //+ " SELECT T1.Pmez, T1.Clave, T1.Ing, T1.Peso / T2.Peso , T1.PesoSeco / T2.PesoSeco "
                            //+ " FROM( "
                            //+ " SELECT rac_descripcion AS Pmez, ing_clave AS Clave, ing_descripcion AS Ing, SUM(rac_mh) AS Peso , SUM(rac_ms) AS PesoSeco "
                            //+ " FROM racion "
                            //+ " WHERE rac_fecha >= '" + fpmI.AddDays(1).ToString("yyyy-MM-dd HH:mm") + "' AND rac_fecha < '" + fin.ToString("yyyy-MM-dd HH:mm") + "' "
                            //+ " AND rac_descripcion LIKE '" + premezcla + "' "
                            //+ " GROUP BY rac_descripcion, ing_clave, ing_descripcion)T1 "
                            //+ " LEFT JOIN( "
                            //+ " SELECT rac_descripcion AS Pmez, SUM(rac_mh) AS Peso , SUM(rac_ms) AS PesoSeco"
                            //+ " FROM racion "
                            //+ " WHERE rac_fecha >= '" + fpmI.AddDays(1).ToString("yyyy-MM-dd HH:mm") + "' AND rac_fecha < '" + fin.ToString("yyyy-MM-dd HH:mm") + "' "
                            //+ " AND rac_descripcion LIKE '" + premezcla + "' "
                            //+ " GROUP BY rac_descripcion"
                            //+ " ) T2 ON T1.Pmez = T2.Pmez";
                            conn.InsertSelecttAlimento(query);
                        }
                    }
                }
            }
            catch (DbException ex) { MessageBox.Show(ex.Message, "ERROR!", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            catch (Exception ex) { MessageBox.Show(ex.Message, "ERROR!", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void SupraMezcla(string premezcla, DateTime inicio, DateTime fin)
        {
            DataTable dt;
            DateTime fini = new DateTime(), ffin;
            DataTable dtF = new DataTable();
            string query = "SELECT * FROM porcentaje_Premezcla where pmez_descripcion like '" + premezcla + "'";
            conn.QueryAlimento(query, out dt);
            int temp = 0;
            DataTable dtV;
            int repeticiones = 0;
            if (dt.Rows.Count == 0)
            {
                query = "SELECT * FROM premezcla WHERE pmez_racion like '" + premezcla + "' AND pmez_fecha <= '" + fin.ToString("yyyy-MM-dd HH:mm") + "'";
                conn.QueryAlimento(query, out dtV);
                if (dtV.Rows.Count > 0)
                {
                    while (dtF.Rows.Count == 0)
                    {
                        if (repeticiones == 30)
                            break;

                        fini = inicio.AddDays(temp);
                        ffin = fini.AddDays(1);

                        query = "SELECT * "
                            + " FROM premezcla "
                            + " WHERE pmez_fecha >= '" + fini.ToString("yyyy-MM-dd HH:mm") + "' AND pmez_fecha< '" + ffin.ToString("yyyy-MM-dd HH:mm") + "' "
                            + " AND pmez_racion like '" + premezcla + "'";
                        conn.QueryAlimento(query, out dtF);
                        temp--;
                        repeticiones++;
                    }

                    query = "INSERT INTO porcentaje_Premezcla "
                    + " SELECT T1.Pmz, T1.Clave, T1.Ing, IIF(T2.Peso > 0 ,(T1.Peso / T2.Peso), 0) , IIF(T3.PORCPESOSECO > 0,(T1.Peso * (T3.PORCPESOSECO))/IIF(T2.Peso > 0, T2.Peso, 0),0) "
                    + " FROM( "
                    + " SELECT pmez_racion AS Pmz, ing_clave AS Clave, ing_nombre AS Ing, SUM(pmez_peso) AS Peso "
                    + " FROM premezcla "
                    + " WHERE pmez_racion LIKE '" + premezcla + "' AND pmez_fecha >= '" + fini.ToString("yyyy-MM-dd HH:mm") + "' AND pmez_fecha < '" + fin.ToString("yyyy-MM-dd HH:mm") + "' "
                    + " GROUP BY pmez_racion, ing_clave, ing_nombre) T1 "
                    + " LEFT JOIN( "
                    + " SELECT pmez_racion AS Pmz, SUM(pmez_peso) AS Peso "
                    + " FROM premezcla "
                    + " WHERE pmez_racion LIKE '" + premezcla + "' AND pmez_fecha >= '" + fini.ToString("yyyy-MM-dd HH:mm") + "' AND pmez_fecha < '" + fin.ToString("yyyy-MM-dd HH:mm") + "' "
                    + " GROUP BY pmez_racion )T2 ON T1.Pmz = T2.Pmz" + @"
                        LEFT JOIN(
	                        SELECT  pmez_racion    AS Pmz
	                               ,ing_clave      AS Clave
	                               ,ing_nombre     AS Ing
	                               ,IIF(Sum(pmez_peso)>0,SUM(pmez_peso * pmez_porcentaje_ms)/Sum(pmez_peso)/100,0) AS PORCPESOSECO
	                        FROM premezcla
	                        WHERE pmez_racion LIKE '" + premezcla + @"'
                            AND pmez_fecha >= '" + inicio.ToString("yyyy-MM-dd HH:mm") + @"'
	                        AND pmez_fecha < '" + fin.ToString("yyyy-MM-dd HH:mm") + @"'
	                        GROUP BY  pmez_racion
	                                 ,ing_clave
	                                 ,ing_nombre
                        ) T3
                        ON T1.Ing = T3.Ing
                        LEFT JOIN(
	                        SELECT  pmez_racion    AS Pmz
	                               ,SUM(pmez_peso * pmez_porcentaje_ms)/100 AS PesoSecoTotal
	                        FROM premezcla
	                        WHERE pmez_racion LIKE '" + premezcla + @"'
                            AND pmez_fecha >= '" + inicio.ToString("yyyy-MM-dd HH:mm") + @"'
	                        AND pmez_fecha < '" + fin.ToString("yyyy-MM-dd HH:mm") + @"'
	                        GROUP BY  pmez_racion 
                        ) T4
                        ON T4.Pmz = T1.Pmz";
                    conn.InsertSelecttAlimento(query);
                }
            }
        }

        private void HoraCorte(out DateTime inicio, out DateTime fin)
        {
            inicio = fecha_inicio;
            fin = fecha_fin;
            int horas = 0, hora_t = 0;
            DataTable dt;
            string query = "select PARAMVALUE FROM bedrijf_params where name like 'DSTimeShift'";
            conn.QueryTracker(query, out dt);

            horas = Convert.ToInt32(dt.Rows[0][0]);
            hora_t = 24 + horas;
            //hora_t = hora_t == 24 ? 0 : hora_t > 24 ? horas : hora_t; //Comentar esto
            if (hora_t == 24)
            {
                inicio = inicio.AddHours(0);
                fin = fin.AddDays(1);
            }
            else if (hora_t > 24)
            {
                inicio = inicio.AddHours(horas);
                fin = fin.AddHours(hora_t);
            }
            else
            {
                inicio = inicio.AddHours(horas);
                fin = fin.AddHours(hora_t);
            }

        }
        private int DiferenciaHoras()
        {
            int horas = 0, hora_t = 0;
            DataTable dt;
            string query = "select PARAMVALUE FROM bedrijf_params where name like 'DSTimeShift'";
            conn.QueryTracker(query, out dt);

            horas = Convert.ToInt32(dt.Rows[0][0]);
            hora_t = 24 + horas;

            return hora_t;
        }
        private void Total(DataTable dt, int[] posiciones)
        {
            int[] sumatoria = new int[posiciones.Length];
            int valor;
            for (int j = 0; j < posiciones.Length; j++)
            {
                valor = 0;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    valor += dt.Rows[i][posiciones[j]] != DBNull.Value ? Convert.ToInt32(dt.Rows[i][posiciones[j]]) : 0;
                }
                sumatoria[j] = valor;
            }


            DataRow row = dt.NewRow();
            row[0] = "TOTAL";
            for (int i = 0; i < sumatoria.Length; i++)
            {
                row[posiciones[i]] = sumatoria[i];
            }
            dt.Rows.Add(row);
        }
        private void AddRow(DataTable dt)
        {
            while (dt.Rows.Count != 31)
            {
                if (dt.Rows.Count < 31)
                {
                    int num = dt.Rows.Count + 1;
                    DataRow row = dt.NewRow();
                    row[0] = "NA";
                    dt.Rows.Add(row);
                }
                else
                {
                    dt.Rows.RemoveAt(dt.Rows.Count - 1);
                }
            }

        }
        private void Hoja2(int julianaI, int julianaF, out DataTable dt)
        {
            ColumnasDTH2(out dt);
            DataTable dtA;
            DataTable DtPrecioLeche = new DataTable();
            DtPrecioLeche.Columns.Add("FECHA").DataType = System.Type.GetType("System.Int32");
            DtPrecioLeche.Columns.Add("PRECIO").DataType = System.Type.GetType("System.Int32");

            for (int i = 0; i < 31; i++)
            {
                DtPrecioLeche.Rows.Add();
            }

            string query = "SELECT T.FECHA, T.IJAULAS, T.JAULAS, T.DESTETADAS, T.MH_DI, T.BEC1, T.PMSBEC1, T.MSBEC1, T.PRECMSBEC1, p.BECERRAS1, "
                         + " T.DESTETADAS2, T.MH_DII, T.BEC2, T.PMSBEC2,T.MSBEC2, T.PRECMSBEC2, p.BECERRAS2, T.VAQUILLAS, T.MH_VQ, T.VAQ,T.PMSVAQ,T.MSVAQ, T.PRECMSVAQ, p.PREÑADAS, "
                        + " T.VACASSECAS, T.MH_S, T.PMSSEC, T.MSSEC, T.SASEC, T.MSSSEC, T.PSSEC, T.SEC, T.PRECMSSEC, p.SECAS, T.IRETO, T.MH_R, T.PMSRETO, T.MSRETO, T.SARETO, T.MSSRETO, T.PSRETO, T.RETO, T.PRECMSRETO, p.RETO, "
                        + " T.IXA, T.CXA , IIF(T.IXA > 0, T.CXA / T.IXA * 100, 0) , T.IT, T.IXA - T.CXA,  IIF(T.IXA > 0, (T.IXA - T.CXA) / T.IXA * 100, 0) "
                        + " FROM( "
                        + " SELECT i.FECHA, i.JAULAS AS IJAULAS, m.JAULAS, i.DESTETADAS, m.MH_DI, m.BEC1, m.MS_DI * 100 AS PMSBEC1, m.MH_DI * m.MS_DI AS MSBEC1, "
                        + " IIf((m.MH_DI * m.MS_DI) > 0, M.BEC1 / (m.MH_DI * m.MS_DI), 0) AS PRECMSBEC1, i.DESTETADAS2, m.MH_DII, m.BEC2, m.MS_DII * 100 AS PMSBEC2, m.MH_DII * m.MS_DII AS MSBEC2, "
                        + " IIf((m.MH_DII * m.MS_DII) > 0, M.BEC2 / (m.MH_DII * m.MS_DII), 0) AS PRECMSBEC2, i.VAQUILLAS, m.MH_VQ, m.VAQ, m.MS_VQ * 100 AS PMSVAQ, m.MH_VQ * m.MS_VQ AS MSVAQ, "
                        + " IIf((m.MH_VQ * m.MS_VQ) > 0, m.VAQ / (m.MH_VQ * m.MS_VQ), 0) AS PRECMSVAQ, i.VACASSECAS, m.MH_S, m.MS_S * 100 AS PMSSEC, m.MH_S * m.MS_S AS MSSEC, m.SA_S AS SASEC, "
                        + " (m.MH_S - m.SA_S) * m.MS_S AS MSSSEC, IIf(m.MH_S > 0, m.SA_S / m.MH_S * 100, 0) AS PSSEC,m.SEC, IIf(m.MS_S > 0 AND m.MH_S > 0, m.SEC / (m.MS_S * m.MH_S), 0) AS PRECMSSEC, "
                        + " (i.VQRETO + i.VCRETO) AS IRETO, m.MH_R, m.MS_R * 100 AS PMSRETO, m.MH_R * m.MS_R AS MSRETO, m.SA_R AS SARETO, (m.MH_R - m.SA_R) * m.MS_R AS MSSRETO, "
                        + " IIF(m.MH_R > 0, m.SA_R / m.MH_R * 100, 0) AS PSRETO, m.RETO, IIF(m.MS_R > 0 AND m.MH_R > 0, m.RETO / (m.MS_R * m.MH_R), 0) AS PRECMSRETO, "
                        + " IIF((i.JAULAS + i.DESTETADAS + i.DESTETADAS2 + i.VAQUILLAS + i.VACASSECAS + i.VQRETO + i.VCRETO + m.VACASORDEÑA) > 0, (m.LECPROD) * PRECIOLECHE / (i.JAULAS + i.DESTETADAS + i.DESTETADAS2 + i.VAQUILLAS + i.VACASSECAS + i.VQRETO + i.VCRETO + i.VACASORDEÑA), 0) AS IXA, "
                        + " IIF((i.JAULAS + i.DESTETADAS + i.DESTETADAS2 + i.VAQUILLAS + i.VACASSECAS + i.VQRETO + i.VCRETO + m.VACASORDEÑA) > 0, "
                        + " ((m.JAULAS * i.JAULAS) + (m.BEC1 * i.DESTETADAS) + (m.BEC2 * i.DESTETADAS2) + (m.VAQ * i.VAQUILLAS) + (m.SEC * i.VACASSECAS) + (m.RETO * (i.VQRETO + i.VCRETO)) + (m.COSTO * i.VACASORDEÑA)) / "
                        + " (i.JAULAS + i.DESTETADAS + i.DESTETADAS2 + i.VAQUILLAS + i.VACASSECAS + i.VQRETO + i.VCRETO + m.VACASORDEÑA), 0) AS CXA, "
                        + " (i.JAULAS + i.DESTETADAS + i.DESTETADAS2 + i.VAQUILLAS + i.VACASSECAS + i.VQRETO + i.VCRETO + m.VACASORDEÑA) AS  IT "
                        + " FROM INVENTARIOAFIXDIA AS i LEFT JOIN MPRODUC AS m ON i.FECHA = m.FECHA "
                        + " WHERE i.FECHA Between " + julianaI + " And " + julianaF + ") T "
                        + " LEFT JOIN PRECIOSTEORICOS P ON P.FECHA = T.FECHA "
                        + " ORDER BY T.FECHA ";
            conn.QueryMovsio(query, out dtA);
            // se buscan los valores del precio de la leche respecto a las fechas
            query = "SELECT DAY(CDATE(FECHA)), PRECIOLECHE FROM MPRODUC "
                     + " WHERE FECHA BETWEEN " + julianaI + " AND " + julianaF
                     + " ORDER BY FECHA";
            conn.QueryMovGanado(query, out DtPrecioLeche);

            if (DtPrecioLeche.Rows.Count == 0)
            {
                for (int i = 0; i < 31; i++)
                {
                    DtPrecioLeche.Rows.Add();
                }

            }

            int auxInventario = 0;
            decimal auxMH = 0, auxCostoRacion = 0, auxCostoMs = 0;

            for (int i = 0; i < dtA.Rows.Count; i++)
            {
                for (int j = 0; j < dtA.Columns.Count; j++)
                {

                    switch (j)
                    {
                        case 5:
                            Int32.TryParse(dtA.Rows[i][3].ToString(), out auxInventario);
                            decimal.TryParse(dtA.Rows[i][4].ToString(), out auxMH);
                            decimal.TryParse(dtA.Rows[i][5].ToString(), out auxCostoRacion);

                            if (auxCostoRacion == 0)
                            {
                                if (auxInventario != 0 && auxMH != 0)
                                    continue;
                                else
                                    dtA.Rows[i][j] = DBNull.Value;
                            }
                            break;
                        case 8:
                            Int32.TryParse(dtA.Rows[i][3].ToString(), out auxInventario);
                            decimal.TryParse(dtA.Rows[i][4].ToString(), out auxMH);
                            decimal.TryParse(dtA.Rows[i][8].ToString(), out auxCostoMs);

                            if (auxCostoMs == 0)
                            {
                                if (auxInventario != 0 && auxMH != 0)
                                    continue;
                                else
                                    dtA.Rows[i][j] = DBNull.Value;
                            }
                            break;                       
                        case 12:
                            Int32.TryParse(dtA.Rows[i][10].ToString(), out auxInventario);
                            decimal.TryParse(dtA.Rows[i][11].ToString(), out auxMH);
                            decimal.TryParse(dtA.Rows[i][12].ToString(), out auxCostoRacion);

                            if (auxCostoRacion == 0)
                            {
                                if (auxInventario != 0 && auxMH != 0)
                                    continue;
                                else
                                    dtA.Rows[i][j] = DBNull.Value;
                            }
                            break;
                        case 15:
                            Int32.TryParse(dtA.Rows[i][10].ToString(), out auxInventario);
                            decimal.TryParse(dtA.Rows[i][11].ToString(), out auxMH);
                            decimal.TryParse(dtA.Rows[i][15].ToString(), out auxCostoMs);

                            if (auxCostoMs == 0)
                            {
                                if (auxInventario != 0 && auxMH != 0)
                                    continue;
                                else
                                    dtA.Rows[i][j] = DBNull.Value;
                            }
                            break;
                        case 19:
                            Int32.TryParse(dtA.Rows[i][17].ToString(), out auxInventario);
                            decimal.TryParse(dtA.Rows[i][8].ToString(), out auxMH);
                            decimal.TryParse(dtA.Rows[i][19].ToString(), out auxCostoRacion);

                            if (auxCostoRacion == 0)
                            {
                                if (auxInventario != 0 && auxMH != 0)
                                    continue;
                                else
                                    dtA.Rows[i][j] = DBNull.Value;
                            }
                            break;
                        case 22:
                            Int32.TryParse(dtA.Rows[i][17].ToString(), out auxInventario);
                            decimal.TryParse(dtA.Rows[i][8].ToString(), out auxMH);
                            decimal.TryParse(dtA.Rows[i][22].ToString(), out auxCostoMs);

                            if (auxCostoMs == 0)
                            {
                                if (auxInventario != 0 && auxMH != 0)
                                    continue;
                                else
                                    dtA.Rows[i][j] = DBNull.Value;
                            }
                            break;
                        case 31:
                            Int32.TryParse(dtA.Rows[i][24].ToString(), out auxInventario);
                            decimal.TryParse(dtA.Rows[i][25].ToString(), out auxMH);
                            decimal.TryParse(dtA.Rows[i][31].ToString(), out auxCostoRacion);
                           
                            if (auxCostoRacion == 0)
                            {
                                if (auxInventario != 0 && auxMH != 0)
                                    continue;
                                else
                                    dtA.Rows[i][j] = DBNull.Value;
                            }
                            break;
                        case 32:
                            Int32.TryParse(dtA.Rows[i][24].ToString(), out auxInventario);
                            decimal.TryParse(dtA.Rows[i][25].ToString(), out auxMH);
                            decimal.TryParse(dtA.Rows[i][32].ToString(), out auxCostoMs);

                            if (auxCostoMs == 0)
                            {
                                if (auxInventario != 0 && auxMH != 0)
                                    continue;
                                else
                                    dtA.Rows[i][j] = DBNull.Value;
                            }
                            break;
                        case 41:
                            Int32.TryParse(dtA.Rows[i][34].ToString(), out auxInventario);
                            decimal.TryParse(dtA.Rows[i][35].ToString(), out auxMH);
                            decimal.TryParse(dtA.Rows[i][41].ToString(), out auxCostoRacion);

                            if (auxCostoRacion == 0)
                            {
                                if (auxInventario != 0 && auxMH != 0)
                                    continue;
                                else
                                    dtA.Rows[i][j] = DBNull.Value;
                            }
                            break;
                        case 42:
                            Int32.TryParse(dtA.Rows[i][34].ToString(), out auxInventario);
                            decimal.TryParse(dtA.Rows[i][35].ToString(), out auxMH);
                            decimal.TryParse(dtA.Rows[i][42].ToString(), out auxCostoMs);

                            if (auxCostoMs == 0)
                            {
                                if (auxInventario != 0 && auxMH != 0)
                                    continue;
                                else
                                    dtA.Rows[i][j] = DBNull.Value;
                            }
                            break;
                        default:
                            if (Convert.ToString(dtA.Rows[i][j]) == "0")
                            {
                                dtA.Rows[i][j] = DBNull.Value;
                            }
                            break;
                    }
                }
            }
            DateTime dia;

            for (int i = 0; i < dtA.Rows.Count; i++)
            {

                dia = ConvertToNormal(Convert.ToInt32(dtA.Rows[i][0]));
                DataRow row = dt.NewRow();
                row[0] = dia.Day.ToString();
                for (int j = 1; j < dtA.Columns.Count; j++)
                    row[j] = dtA.Rows[i][j];
                dt.Rows.Add(row);
            }

            DataTable dtInv;
            int jaulas = 0, d1 = 0, d2 = 0, vp = 0, secas = 0, reto = 0;
            query = "SELECT FECHA, JAULAS, DESTETADAS, DESTETADAS2, VAQUILLAS, VACASSECAS, (VQRETO + VCRETO) AS RETO "
                    + " FROM INVENTARIOAFIXDIA "
                    + " WHERE FECHA BETWEEN " + julianaI + " AND " + julianaF;
            conn.QueryMovsio(query, out dtInv);

            for (int i = 0; i < dtInv.Rows.Count; i++)
            {
                if (dtInv.Rows[i][1] != DBNull.Value) { jaulas += Convert.ToInt32(dtInv.Rows[i][1]); }
                if (dtInv.Rows[i][2] != DBNull.Value) { d1 += Convert.ToInt32(dtInv.Rows[i][2]); }
                if (dtInv.Rows[i][3] != DBNull.Value) { d2 += Convert.ToInt32(dtInv.Rows[i][3]); }
                if (dtInv.Rows[i][4] != DBNull.Value) { vp += Convert.ToInt32(dtInv.Rows[i][4]); }
                if (dtInv.Rows[i][5] != DBNull.Value) { secas += Convert.ToInt32(dtInv.Rows[i][5]); }
                if (dtInv.Rows[i][6] != DBNull.Value) { reto += Convert.ToInt32(dtInv.Rows[i][6]); }
            }

            for (int i = dt.Rows.Count; i < 36; i++)
            {
                if (i == 31)
                {
                    DataRow dr = dt.NewRow();
                    dr["INV"] = jaulas;
                    dr["INV2"] = d1;
                    dr["INV7"] = d2;
                    dr["INV13"] = vp;
                    dr["INVSECAS"] = secas;
                    dr["INVRETO"] = reto;
                    dr["DIA"] = "TOTAL";
                    dt.Rows.Add(dr);
                }
                else
                {
                    dt.Rows.Add();
                }
            }
            dt.Rows[32][0] = "PROM";
            dt.Rows[33][0] = "AÑO ANT";
            dt.Rows[34][0] = "DIF %";
            dt.Rows[35][0] = "DIF #";

            double _IT = 0, _LecheXPrecio = 0, _PondeTablas = 0;
            double[] _SumatoriaPreciosxInventario = { 0, 0, 0, 0, 0, 0, 0, 0 };
            //Se sacan las sumatorias de los inventarios 
            if (TablaGlobal.Rows[31][2] == DBNull.Value) { TablaGlobal.Rows[31][2] = 0; }
            if (dt.Rows[31][3] == DBNull.Value) { dt.Rows[31][3] = 0; }
            if (dt.Rows[31][10] == DBNull.Value) { dt.Rows[31][10] = 0; }
            if (dt.Rows[31][17] == DBNull.Value) { dt.Rows[31][17] = 0; }
            if (dt.Rows[31][24] == DBNull.Value) { dt.Rows[31][24] = 0; }
            if (dt.Rows[31][34] == DBNull.Value) { dt.Rows[31][34] = 0; }
            if (dt.Rows[31][1] == DBNull.Value) { dt.Rows[31][1] = 0; }
            _PondeTablas = Convert.ToDouble(TablaGlobal.Rows[31][2]) + Convert.ToDouble(dt.Rows[31][1]) + Convert.ToDouble(dt.Rows[31][3]) + Convert.ToDouble(dt.Rows[31][10]) + Convert.ToDouble(dt.Rows[31][17]) + Convert.ToDouble(dt.Rows[31][24]) + Convert.ToDouble(dt.Rows[31][34]);
            for (int i = 0; i < 31; i++)
            {
                //Se saca el valor de la leche multiplicada por el precio 
                if (TablaGlobal.Rows[i][11] != DBNull.Value && DtPrecioLeche.Rows[i][1] != DBNull.Value)
                {
                    _LecheXPrecio += Convert.ToDouble(TablaGlobal.Rows[i][11]) * Convert.ToDouble(DtPrecioLeche.Rows[i][1]);
                }
                //Sacamos la sumatoria del ordeño por el precio
                if (TablaGlobal.Rows[i][2] != DBNull.Value && TablaGlobal.Rows[i][31] != DBNull.Value)
                {

                    _SumatoriaPreciosxInventario[0] += Convert.ToDouble(Convert.ToDouble(TablaGlobal.Rows[i][2])) * Convert.ToDouble(Convert.ToDouble(TablaGlobal.Rows[i][31]));
                }
            }

            double _Aux1 = 0, _Aux2 = 0;
            for (int i = 1; i < 50; i++)
            {
                int _seguro = 0, _contador = 0;
                double _Ponderizado = 0;
                for (int j = 0; j <= 30; j++)
                {
                    if (dt.Rows[j][i] == DBNull.Value) { /*_contador++; */}
                    else
                    {
                        if (_seguro == 0)
                        {
                            dt.Rows[32][i] = 0;
                            _seguro = 1;
                        }
                        if (i == 1 || i == 3 || i == 10 || i == 17 || i == 24 || i == 34)
                        {
                            dt.Rows[32][i] = Convert.ToInt32(dt.Rows[32][i]) + Convert.ToInt32(dt.Rows[j][i]);
                            _contador++;
                        }
                        else
                        {
                            // Sacamos los valores ponderizados
                            if (dt.Rows[j][1] != DBNull.Value && i == 2)
                            {
                                _Ponderizado += Convert.ToDouble(dt.Rows[j][i]) * Convert.ToDouble(dt.Rows[j][1]);
                            }
                            if (dt.Rows[j][3] != DBNull.Value && i >= 4 && i <= 8)
                            {
                                _Ponderizado += Convert.ToDouble(dt.Rows[j][i]) * Convert.ToDouble(dt.Rows[j][3]);
                            }
                            if (dt.Rows[j][10] != DBNull.Value && i >= 11 && i <= 15)
                            {
                                _Ponderizado += Convert.ToDouble(dt.Rows[j][i]) * Convert.ToDouble(dt.Rows[j][10]);
                            }
                            if (dt.Rows[j][17] != DBNull.Value && i >= 18 && i <= 22)
                            {
                                _Ponderizado += Convert.ToDouble(dt.Rows[j][i]) * Convert.ToDouble(dt.Rows[j][17]);
                            }
                            if (dt.Rows[j][24] != DBNull.Value && i >= 25 && i <= 32)
                            {
                                _Ponderizado += Convert.ToDouble(dt.Rows[j][i]) * Convert.ToDouble(dt.Rows[j][24]);
                            }
                            if (dt.Rows[j][34] != DBNull.Value && i >= 35 && i <= 42)
                            {
                                _Ponderizado += Convert.ToDouble(dt.Rows[j][i]) * Convert.ToDouble(dt.Rows[j][34]);
                            }
                            //Si no es necesario sacar el valor ponderizado simplimente se saca el valor promedio 
                            dt.Rows[32][i] = Convert.ToDouble(dt.Rows[32][i]) + Convert.ToDouble(dt.Rows[j][i]);
                            _contador++;
                        }
                    }
                }
                if (dt.Rows[32][i] == DBNull.Value) { }
                else
                {
                    if (i == 1 || i == 3 || i == 10 || i == 17 || i == 24 || i == 34)
                    {
                        //dt.Rows[32][i] = Decimal.Round(Convert.ToInt32(dt.Rows[32][i]) / _contador);
                        dt.Rows[32][i] = Math.Round(Convert.ToDouble(dt.Rows[32][i]) / _contador);
                    }
                    else
                    {
                        if (i == 47)
                        {
                            //Tomamos la sumatoria del valor IT 
                            _IT = Convert.ToDouble(dt.Rows[32][47]);
                        }
                        dt.Rows[32][i] = Convert.ToDouble(dt.Rows[32][i]) / _contador;

                        if (i == 2)
                        {
                            _SumatoriaPreciosxInventario[7] = _Ponderizado;
                            dt.Rows[32][i] = _Ponderizado / Convert.ToDouble(dt.Rows[31][1]);
                        }
                        if (i >= 4 && i <= 8)
                        {
                            //Guardamos el valor de la sumatoria del precio 2/7
                            if (i == 5) { _SumatoriaPreciosxInventario[1] = _Ponderizado; }
                            if (i == 7) { dt.Rows[32][6] = ((_Ponderizado / Convert.ToDouble(dt.Rows[31][3])) / Convert.ToDouble(dt.Rows[32][4])) * 100; }
                            dt.Rows[32][i] = _Ponderizado / Convert.ToDouble(dt.Rows[31][3]);
                            if (i == 8) { dt.Rows[32][i] = Convert.ToDouble(dt.Rows[32][5]) / Convert.ToDouble(dt.Rows[32][7]); }
                        }
                        if (i >= 11 && i <= 15)
                        {
                            //Guardamos el valor de la sumatoria del precio 2/13
                            if (i == 12) { _SumatoriaPreciosxInventario[2] = _Ponderizado; }
                            if (i == 14)
                            {
                                dt.Rows[32][13] = ((_Ponderizado / Convert.ToDouble(dt.Rows[31][10])) / Convert.ToDouble(dt.Rows[32][11])) * 100;
                            }
                            dt.Rows[32][i] = _Ponderizado / Convert.ToDouble(dt.Rows[31][10]);
                            if (i == 15) { dt.Rows[32][i] = Convert.ToDouble(dt.Rows[32][12]) / Convert.ToDouble(dt.Rows[32][14]); }
                        }
                        if (i >= 18 && i <= 22)
                        {
                            //Guardamos el valor de la sumatoria del precio de 13 o más 
                            if (i == 19) { _SumatoriaPreciosxInventario[3] = _Ponderizado; }
                            if (i == 21)
                            {
                                dt.Rows[32][19] = ((_Ponderizado / Convert.ToDouble(dt.Rows[31][17])) / Convert.ToDouble(dt.Rows[32][18])) * 100;
                            }
                            dt.Rows[32][i] = _Ponderizado / Convert.ToDouble(dt.Rows[31][17]);
                            if (i == 22) { dt.Rows[32][i] = Convert.ToDouble(dt.Rows[32][20]) / Convert.ToDouble(dt.Rows[32][21]); }
                        }
                        if (i >= 25 && i <= 32)
                        {
                            //Guardamos el valor de la sumatoria de secas
                            if (i == 31) { _SumatoriaPreciosxInventario[4] = _Ponderizado; }
                            dt.Rows[32][i] = _Ponderizado / Convert.ToDouble(dt.Rows[31][24]);
                            if (i == 27) { dt.Rows[32][26] = ((_Ponderizado / Convert.ToDouble(dt.Rows[31][24])) / Convert.ToDouble(dt.Rows[32][25])) * 100; }
                            if (i == 25) { _Aux1 = Convert.ToDouble(dt.Rows[32][i]); }
                            if (i == 28) { _Aux2 = Convert.ToDouble(dt.Rows[32][i]); }
                            if (i == 30) { dt.Rows[32][i] = Convert.ToDouble(dt.Rows[32][28]) / Convert.ToDouble(dt.Rows[32][25]); }
                            if (i == 32) { dt.Rows[32][i] = Convert.ToDouble(dt.Rows[32][31]) / Convert.ToDouble(dt.Rows[32][29]); }
                        }

                        if (i >= 35 && i <= 42)
                        {
                            //Guardamos el valor de la sumatoria del reto
                            if (i == 41) { _SumatoriaPreciosxInventario[5] = _Ponderizado; }
                            dt.Rows[32][i] = _Ponderizado / Convert.ToDouble(dt.Rows[31][34]);
                            if (i == 37) { dt.Rows[32][36] = ((_Ponderizado / Convert.ToDouble(dt.Rows[31][34])) / Convert.ToDouble(dt.Rows[32][35])) * 100; }
                            if (i == 35) { _Aux1 = Convert.ToDouble(dt.Rows[32][i]); }
                            if (i == 38) { _Aux2 = Convert.ToDouble(dt.Rows[32][i]); }
                            if (i == 40) { dt.Rows[32][i] = Convert.ToDouble(dt.Rows[32][38]) / Convert.ToDouble(dt.Rows[32][35]); }
                            if (i == 42) { dt.Rows[32][i] = Convert.ToDouble(dt.Rows[32][41]) / Convert.ToDouble(dt.Rows[32][39]); }
                        }
                        if (i == 30 || i == 40)
                        {
                            dt.Rows[32][i] = (_Aux2 / _Aux1) * 100;
                        }
                    }
                }
            }
            //Sacamos la sumatoria de todas las sumatorias de precios por invenarios y ordeño por precio
            _SumatoriaPreciosxInventario[6] = _SumatoriaPreciosxInventario[5] + _SumatoriaPreciosxInventario[4] + _SumatoriaPreciosxInventario[3] +
                                              _SumatoriaPreciosxInventario[2] + _SumatoriaPreciosxInventario[1] + _SumatoriaPreciosxInventario[0] +
                                              _SumatoriaPreciosxInventario[7];
            dt.Rows[32][44] = _LecheXPrecio / _IT;
            dt.Rows[32][45] = _SumatoriaPreciosxInventario[6] / _PondeTablas;
            dt.Rows[32][46] = Convert.ToDouble(dt.Rows[32][44]) > 0 ? (Convert.ToDouble(dt.Rows[32][45]) / Convert.ToDouble(dt.Rows[32][44])) * 100 : 0;
            dt.Rows[32][48] = Convert.ToDouble(dt.Rows[32][44]) - Convert.ToDouble(dt.Rows[32][45]);
            dt.Rows[32][49] = Convert.ToDouble(dt.Rows[32][44]) > 0 ? (Convert.ToDouble(dt.Rows[32][48]) / Convert.ToDouble(dt.Rows[32][44])) * 100 : 0;

            DateTime fecha = ConvertToNormal(julianaF);
            if (fecha.Year == fechaFinDTP.Year)
            {
                utilidad.IXA = Convert.ToDecimal(dt.Rows[32][44]);
                utilidad.CXA = Convert.ToDecimal(dt.Rows[32][45]);
                utilidad.PORCENTAJE_C = Convert.ToDecimal(dt.Rows[32][46]);
                utilidad.UXA = Convert.ToDecimal(dt.Rows[32][48]);
                utilidad.PORCENTAJE_U = Convert.ToDecimal(dt.Rows[32][49]);
            }
        }

        private void ColumnasDTH1(out DataTable dt)
        {
            dt = new DataTable();
            dt.Columns.Add("DIA").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("TOTALVTA").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("ORDENO").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("SECAS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("HATO").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCLACT").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCPROT").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("UREA").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCGRA").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("CCS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("CTD").DataType = System.Type.GetType("System.Double");
            //PROD
            dt.Columns.Add("LECHE").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("ANTIB").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("X").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("TOTALPROD").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("DEL").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("ANT").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("METAPROD").DataType = System.Type.GetType("System.Double");
            //VENTA
            dt.Columns.Add("ILCAVTA").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("ICVTA").DataType = System.Type.GetType("System.Double");
            //ALIMENTACION PRODUCCION
            dt.Columns.Add("EA").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("ILCAPROD").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("ICPROD").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIOL").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("MH").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCMS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("MS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("SA").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("MSS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("EAS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIOPROD").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIOMS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("METAMS").DataType = System.Type.GetType("System.Double");
            //CRIBAS
            dt.Columns.Add("N1").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("N2").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("N3").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("N4").DataType = System.Type.GetType("System.Double");
            //NO ID
            dt.Columns.Add("SES1").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("SES2").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("SES3").DataType = System.Type.GetType("System.Int32");
            //DEC 
            dt.Columns.Add("DEC").DataType = System.Type.GetType("System.Double");
        }
        private void ColumnasDTH2(out DataTable dt)
        {
            dt = new DataTable();
            dt.Columns.Add("DIA").DataType = System.Type.GetType("System.String");
            //Jaulas
            dt.Columns.Add("INV").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("PRECIOJ").DataType = System.Type.GetType("System.Double");
            // 2/7
            dt.Columns.Add("INV2").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("MH2").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIO2").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCMS2").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("MS2").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIOMS2").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("METABEC1").DataType = System.Type.GetType("System.Double");
            // 7/13
            dt.Columns.Add("INV7").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("MH7").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIO7").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCMS7").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("MS7").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIOMS7").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("METABEC2").DataType = System.Type.GetType("System.Double");
            // 13 a Mas
            dt.Columns.Add("INV13").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("MH13").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIO13").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCMS13").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("MS13").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIOMS13").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("METAVP").DataType = System.Type.GetType("System.Double");
            //Secas
            dt.Columns.Add("INVSECAS").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("MHSECAS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCMSSECAS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("MSSECAS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("SASECAS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("MSSSECAS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCSSECAS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIOSECAS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIOMSSECAS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("METASECAS").DataType = System.Type.GetType("System.Double");
            //Reto
            dt.Columns.Add("INVRETO").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("MHRETO").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCMSRETO").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("MSRETO").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("SARETO").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("MSSRETO").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCSRETO").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIORETO").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIOMSRETO").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("METARETO").DataType = System.Type.GetType("System.Double");
            //Utilidad Por Animal
            dt.Columns.Add("IXA").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("CXA").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCENTAJE1").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("IT").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("UXA").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCENTAJE2").DataType = System.Type.GetType("System.Double");
        }
        private DateTime ConvertToNormal(int Normal)
        {
            DateTime dt = new DateTime(1900, 01, 01);
            try
            {
                dt = dt.AddDays(Normal);
                dt = dt.AddDays(-2);
                return dt;
            }
            catch
            {
                dt = DateTime.Today; ;
                MessageBox.Show("Error en conversión de fecha");
                return dt;
            }
        }
        private void InfoDPA(out DataTable _DtAux)
        {
            fecha_fin = monthCalendar1.SelectionRange.End;
            fecha_inicio = new DateTime(fecha_fin.Year, fecha_fin.Month, 1);
            //Sacamos la fecha en juliana correspondiente
            int _FechaFinalEnJuliana = ConvertToJulian(fecha_fin);
            int _FechaInicialEnJuliana = ConvertToJulian(fecha_inicio);

            //Sacamos la fecha en juliana correspondiente al año anterior 
            int _FechaFinalEnJulianaAñoAnt = ConvertToJulian(fecha_fin.AddYears(-1));
            int _FechaInicialEnJulianaAñoAnt = ConvertToJulian(fecha_inicio.AddYears(-1));

            ColumnasDPA(out _DtAux);
            //Esto nos sirve para saber el día que sea solo dando las fechas en julianas (era para hacer pruebas)
            //_FechaInicialEnJuliana = 44256;
            //_FechaFinalEnJuliana = _FechaFinalEnJuliana - 1; ;//44286; 44256
            //_FechaInicialEnJulianaAñoAnt = 43891;
            //_FechaFinalEnJulianaAñoAnt = _FechaFinalEnJulianaAñoAnt - 1; //43921; //43921;43891

            for (int i = 0; i < 35; i++)
            {
                DataRow _Rows = _DtAux.NewRow();
                _DtAux.Rows.Add(_Rows);
            }
            _DtAux.Rows[31][0] = "TOTAL";
            _DtAux.Rows[32][0] = "AÑO ANT";
            _DtAux.Rows[33][0] = "DIF %";
            _DtAux.Rows[34][0] = "DIF #";
            DataTable _dt;
            //Consulta para saber jaulas vivas
            string query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 2  AND motivo <> 1 AND destino <> 2 "
                + " AND (edovac = 1 OR edovac = 10)"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            //La primera vez que se manda a llamar LlenarEspaciosVacios se manda un true con el objetivo de
            //sacar las primeras dos columnas incluida la de la fecha, despues es un flase para solo llenar 
            //la columna que nos interesa
            LlenarEspaciosVacios(_DtAux, _dt, true, 1);

            //Con este FOR se llena de espacios vacios los días que sean despues de la fecha actual
            for (int i = fecha_fin.Day; i <= 30; i++)
            {
                _DtAux.Rows[i][0] = "NA";
            }

            //Consulta para saber jaulas muertas 
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 2  AND motivo = 1 "
                + " AND (edovac = 1 OR edovac = 10)"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 2);

            //Consulta para saber destete vivas
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 2  AND motivo <> 1 AND destino <> 2"
                + " AND (edovac = 2 OR edovac = 7)"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 3);

            //Consulta para saber destete muertas 
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 2  AND motivo = 1"
                + " AND (edovac = 2 OR edovac = 7)"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 4);

            //Consulta para saber vaquillas muertas 
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 2  AND motivo = 1 AND destino = 1"
                + " AND (edovac = 3 OR edovac = 8 OR edovac = 9)"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 5);

            //Consulta para saber vaquillas en urgencia  
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 2  AND motivo = 2 AND destino = 1"
                + " AND (edovac = 3 OR edovac = 8 OR edovac = 9)"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 6);

            //Consulta para saber vaquillas delgadas 
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 2  AND motivo = 3 AND destino = 1"
                + " AND (edovac = 3 OR edovac = 8 OR edovac = 9)"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 7);

            //Consulta para saber vaquillas regulares
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 2  AND motivo = 4 AND destino = 1"
                + " AND (edovac = 3 OR edovac = 8 OR edovac = 9)"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 8);

            //Consulta para saber vaquillas gordas
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 2  AND motivo = 5 AND destino = 1"
                + " AND (edovac = 3 OR edovac = 8 OR edovac = 9)"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 9);

            //Consulta para saber vaquillas otros casos
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 2  AND motivo > 5 AND destino = 1"
                + " AND (edovac = 3 OR edovac = 8 OR edovac = 9)"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 10);

            //Consulta para saber vacas muertas 
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 1  AND motivo = 1 AND destino = 1"
                + " AND (edovac > 3)"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 11);

            //Consulta para saber vacas en urgencia  
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 1  AND motivo = 2 AND destino = 1"
                + " AND (edovac > 3)"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 12);

            //Consulta para saber vacas delgadas
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 1  AND motivo = 3 AND destino = 1"
                + " AND (edovac > 3)"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 13);

            //Consulta para saber vacas regulares
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 1  AND motivo = 4 AND destino = 1"
                + " AND (edovac > 3)"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 14);

            //Consulta para saber vacas gordas
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 1  AND motivo = 5 AND destino = 1"
                + " AND (edovac > 3)"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 15);

            //Consulta para saber vacas otros casos
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 1  AND motivo > 5 AND destino = 1"
                + " AND (edovac > 3)"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 16);

            //Consulta para saber vaquillas N/D
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM (SELECT distinct FECHA, numvac FROM Dnacimiento "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 2  AND (tipopar = 1 or tipopar = 3))"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 17);

            //Consulta para saber abortos de vaquillas
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM (SELECT distinct FECHA, numvac FROM Dnacimiento "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 2  AND (tipopar = 2))"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 18);

            //Consulta para saber nacimientos de vaquillas hembras 
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM Dnacimiento "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 2  AND criaviva = 1  AND criasexo = 1"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 19);

            //Consulta para saber nacimientos de vaquillas machos
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM Dnacimiento "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 2  AND criaviva = 1  AND criasexo = 2"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 20);

            //Consulta para saber nacimientos de vacas N/D
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM (SELECT distinct FECHA, numvac FROM Dnacimiento "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 1  AND (tipopar = 1 or tipopar = 3))"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 21);

            //Consulta para saber abortos de vacas
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM (SELECT distinct FECHA, numvac FROM Dnacimiento "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 1  AND (tipopar = 2))"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 22);

            //Consulta para saber nacimientos de vacas hembras 
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM Dnacimiento "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 1  AND criaviva = 1  AND criasexo = 1"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 23);

            //Consulta para saber nacimientos de vacas machos 
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM Dnacimiento "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 1  AND criaviva = 1  AND criasexo = 2"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 24);

            //Tomamos los valores de los partos sin crias
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM (SELECT distinct FECHA, numvac FROM Dnacimiento "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND tipopar <> 2 AND criaviva = 2  AND criasexo = 3)"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 25);

            //Tomamos los valores de los partos vacas
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM (SELECT distinct FECHA, numvac FROM Dnacimiento "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 2)"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 26);

            //Tomamos los valores de los partos vaquillas
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM (SELECT distinct FECHA, numvac FROM Dnacimiento "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 1)"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 27);

            //Consulta para saber las muertas al nacer de día
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*) "
                + " FROM Dnacimiento "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND criaviva = 2 AND (tipopar = 1 or tipopar = 3) AND criasexo < 3 AND dianoche = 1 "
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 28);

            //Consulta para saber las muertas al nacer de noche
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*) "
                + " FROM Dnacimiento "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND criaviva = 2 AND (tipopar = 1 or tipopar = 3) AND criasexo < 3 AND dianoche = 2 "
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 29);

            //Tomamos los valores de los abortos de vaquillas 
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM Dabortos "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 2"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 30);

            //Tomamos los valores de los abortos de vaquillas 
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                + " FROM Dabortos "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                + " AND vacvaq = 1"
                + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dt);
            SacarEspaciosVacios(_dt);
            LlenarEspaciosVacios(_DtAux, _dt, false, 31);

            //Consulta para saber jaulas vivas del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 2  AND motivo <> 1 AND destino <> 2 "
                + " AND (edovac = 1 OR edovac = 10)"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][1] = _dt.Rows[0][0];

            //Consulta para saber jaulas muertas del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 2  AND motivo = 1 "
                + " AND (edovac = 1 OR edovac = 10)"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][2] = _dt.Rows[0][0];

            //Consulta para saber destete vivas del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 2  AND motivo <> 1 AND destino <> 2 "
                + " AND (edovac = 2 OR edovac = 7)"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][3] = _dt.Rows[0][0];

            //Consulta para saber destete muertas  del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 2  AND motivo = 1"
                + " AND (edovac = 2 OR edovac = 7)"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][4] = _dt.Rows[0][0];

            //Consulta para saber vaquillas muertas del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 2  AND motivo = 1 AND destino = 1"
                + " AND (edovac = 3 OR edovac = 8 OR edovac = 9)"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][5] = _dt.Rows[0][0];

            //Consulta para saber vaquillas en urgencia del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 2  AND motivo = 2 AND destino = 1"
                + " AND (edovac = 3 OR edovac = 8 OR edovac = 9)"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][6] = _dt.Rows[0][0];

            //Consulta para saber vaquillas en urgencia del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 2  AND motivo = 3 AND destino = 1"
                + " AND (edovac = 3 OR edovac = 8 OR edovac = 9)"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][7] = _dt.Rows[0][0];

            //Consulta para saber vaquillas regulares del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 2  AND motivo = 4 AND destino = 1"
                + " AND (edovac = 3 OR edovac = 8 OR edovac = 9)"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][8] = _dt.Rows[0][0];

            //Consulta para saber vaquillas gordas del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 2  AND motivo = 5 AND destino = 1"
                + " AND (edovac = 3 OR edovac = 8 OR edovac = 9)"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][9] = _dt.Rows[0][0];

            //Consulta para saber vaquillas otros casos del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 2  AND motivo > 5 AND destino = 1"
                + " AND (edovac = 3 OR edovac = 8 OR edovac = 9)"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][10] = _dt.Rows[0][0];

            //Consulta para saber vacas muertas del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 1  AND motivo = 1 AND destino = 1 "
                + " AND (edovac > 3)"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][11] = _dt.Rows[0][0];

            //Consulta para saber vacas en urgencia del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 1  AND motivo = 2 AND destino = 1 "
                + " AND (edovac > 3)"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][12] = _dt.Rows[0][0];

            //Consulta para saber vacas delgadas del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 1  AND motivo = 3 AND destino = 1 "
                + " AND (edovac > 3)"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][13] = _dt.Rows[0][0];

            //Consulta para saber vacas regulares del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 1  AND motivo = 4 AND destino = 1 "
                + " AND (edovac > 3)"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][14] = _dt.Rows[0][0];

            //Consulta para saber vacas gordas del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 1  AND motivo = 5 AND destino = 1 "
                + " AND (edovac > 3)"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][15] = _dt.Rows[0][0];

            //Consulta para saber vacas otros casos del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM Ddesecho "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 1  AND motivo > 5 AND destino = 1 "
                + " AND (edovac > 3)"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][16] = _dt.Rows[0][0];

            //Consulta para saber saber vaquillas N/D del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM (SELECT distinct FECHA, numvac FROM Dnacimiento "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 2  AND (tipopar = 1 or tipopar = 3)) "
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][17] = _dt.Rows[0][0];

            //Consulta para saber saber abortos de vaquillas del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM (SELECT distinct FECHA, numvac FROM Dnacimiento "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 2  AND (tipopar = 2)) "
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][18] = _dt.Rows[0][0];

            //Consulta para saber saber nacimientos de vaquillas hembras del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM Dnacimiento "
                + "WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 2  AND criaviva = 1  AND criasexo = 1"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][19] = _dt.Rows[0][0];

            //Consulta para saber saber nacimientos de vaquillas machos del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM Dnacimiento "
                + "WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 2  AND criaviva = 1  AND criasexo = 2"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][20] = _dt.Rows[0][0];

            //Consulta para saber saber nacimientos de vacas N/D año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM (SELECT distinct FECHA, numvac FROM Dnacimiento "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 1  AND (tipopar = 1 or tipopar = 3)) "
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][21] = _dt.Rows[0][0];

            //Consulta para saber saber abortos de vacas del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM (SELECT distinct FECHA, numvac FROM Dnacimiento "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 1  AND (tipopar = 2)) "
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][22] = _dt.Rows[0][0];

            //Consulta para saber saber nacimientos de vacas hembras del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM Dnacimiento "
                + "WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 1  AND criaviva = 1  AND criasexo = 1"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][23] = _dt.Rows[0][0];

            //Consulta para saber saber nacimientos de vacas machos del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM Dnacimiento "
                + "WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 1  AND criaviva = 1  AND criasexo = 2"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][24] = _dt.Rows[0][0];

            //Tomamos los valores de los partos sin crias del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM (SELECT distinct FECHA, numvac FROM Dnacimiento "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND tipopar <> 2 AND criaviva = 2  AND criasexo = 3)"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][25] = _dt.Rows[0][0];

            //Tomamos los valores de los partos vacas del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM (SELECT distinct FECHA, numvac FROM Dnacimiento "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 2)"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][26] = _dt.Rows[0][0];

            //Tomamos los valores de los partos vaquillas del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM (SELECT distinct FECHA, numvac FROM Dnacimiento "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 1)"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][27] = _dt.Rows[0][0];

            //Tomamos las muertas al nacer de día del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM Dnacimiento "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND criaviva = 2 AND (tipopar = 1 or tipopar = 3) AND criasexo < 3 AND dianoche = 1"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][28] = _dt.Rows[0][0];

            //Tomamos las muertas al nacer de día del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM Dnacimiento "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND criaviva = 2 AND (tipopar = 1 or tipopar = 3) AND criasexo < 3 AND dianoche = 2"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][29] = _dt.Rows[0][0];

            //Tomamos las muertas al nacer de día del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM Dabortos "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 2"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][30] = _dt.Rows[0][0];

            //Tomamos las muertas al nacer de día del año anterior 
            query = "SELECT  SUM (t.NUMVACAS) FROM"
                + " (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS "
                + " FROM Dabortos "
                + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
                + " AND vacvaq = 1"
                + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dt);
            _DtAux.Rows[32][31] = _dt.Rows[0][0];

            //Este for sirve para el llenado de los ultimos dos renglones, validando los valores que hay y aplicando las formulas correspondientes 
            for (int i = 1; i < _DtAux.Columns.Count; i++)
            {
                if (_DtAux.Rows[31][i] == DBNull.Value)
                {
                }
                else
                {
                    if (_DtAux.Rows[32][i] == DBNull.Value)
                    {
                        _DtAux.Rows[33][i] = 100;
                        _DtAux.Rows[34][i] = Convert.ToInt32(_DtAux.Rows[31][i]);
                    }
                    else
                    {
                        //_DtAux.Rows[33][i] = Decimal.Round((((Convert.ToDecimal(_DtAux.Rows[32][i]) / Convert.ToDecimal(_DtAux.Rows[31][i])) - 1) * -100));
                        //_DtAux.Rows[33][i] = Math.Round((((Convert.ToDecimal(_DtAux.Rows[32][i]) / Convert.ToDecimal(_DtAux.Rows[31][i])) - 1) * -100));
                        _DtAux.Rows[34][i] = Convert.ToInt32(_DtAux.Rows[31][i]) - Convert.ToInt32(_DtAux.Rows[32][i]);
                        _DtAux.Rows[33][i] = Math.Round((((Convert.ToDecimal(_DtAux.Rows[34][i]) / Convert.ToDecimal(_DtAux.Rows[32][i]))) * 100));
                        if (Convert.ToDouble(_DtAux.Rows[34][i]) > 0)
                        {
                            _DtAux.Rows[33][i] = Math.Abs(Convert.ToDouble(_DtAux.Rows[33][i]));
                        }

                    }
                }
            }

        }

        private void LlenarEspaciosVacios(DataTable _DtAux, DataTable _dt, bool _dia, int _columnas)
        {
            //Se crea la variable seguro ya que se busca que un proceso se haga solo una vez siempre y cuando se cumpla la condicion de haber un valor que sumar 
            int _seguro = 0;
            //En esta parte asignamos un valor para no tener problemas con nulos en el total correspondiente a la columna
            //Este FOR nos ayuda a revisar celda por celda el valor que tenemos en el datatable 
            for (int i = 0; i < _dt.Rows.Count; i++)
            {
                //Supongamos que encontramos un valor en el datatable por lo que hay que sumar el valor actual en total y el
                if (_dt.Rows[i][1] != DBNull.Value && _seguro == 0)
                {
                    _DtAux.Rows[31][_columnas] = 0;
                    _seguro = 1;
                }
                //En caso de tener un valor nulo no hace nada de otra manera entra a la celda y suma el valor con el valo actual en la celda total
                if (_dt.Rows[i][1] != DBNull.Value)
                {
                    _DtAux.Rows[31][_columnas] = Convert.ToInt32(_DtAux.Rows[31][_columnas]) + Convert.ToInt32(_dt.Rows[i][1]);
                }
            }
            // Este if nos permite agregar los valores a la columna día y a la segunda columna 
            if (_dia)
            {
                for (int i = 0; i < _dt.Rows.Count; i++)
                {
                    _DtAux.Rows[i][0] = _dt.Rows[i][0];
                    _DtAux.Rows[i][_columnas] = _dt.Rows[i][1];
                }
            }
            // En caso contrario que sea la tercera
            // o más simplemente agrega los valores de la columna 
            else
            {
                for (int i = 0; i < _dt.Rows.Count; i++)
                {
                    _DtAux.Rows[i][_columnas] = _dt.Rows[i][1];
                }
            }
        }

        private void SacarEspaciosVacios(DataTable _dt)
        {

            int _dias, _index;
            ArrayList _EspacioVacios = new ArrayList();
            for (int i = 0; i < 31; i++)
            {
                _EspacioVacios.Add(i + 1);
            }
            for (int i = 0; i < _dt.Rows.Count; i++)
            {
                _dias = Convert.ToInt32(_dt.Rows[i][0]);
                _index = _EspacioVacios.IndexOf(_dias);
                _EspacioVacios.RemoveAt(_index);
            }
            for (int i = 0; i < _EspacioVacios.Count; i++)
            {
                DataRow _row = _dt.NewRow();
                _index = Convert.ToInt32(_EspacioVacios[i]);
                _row[0] = _index;
                _dt.Rows.InsertAt(_row, _index - 1);
            }

        }

        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {

        }

        private void ColumnasDPA(out DataTable dt)
        {
            dt = new DataTable();
            dt.Columns.Add("DIA").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("J_VIVAS").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("J_MTAS").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("D_VIVAS").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("D_MTAS").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VQ_MTAS").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VQ_URG").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VQ_RF").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VQ_RR").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VQ_RG").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VQ_OTROS").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VC_MTAS").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VC_URG").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VC_RF").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VC_RR").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VC_RG").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VC_OTROS").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VQ_ND").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VQ_A").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VQ_H").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VQ_M").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VC_ND").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VC_A").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VC_H").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VC_M").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("SC").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("TP_VQ").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("TP_VC").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("MTAS_DIA").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("MTAS_NOC").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("AB_VQ").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("AB_VC").DataType = System.Type.GetType("System.Int32");

        }

        private void InfoSR(out DataTable _DtSR)
        {
            fecha_fin = monthCalendar1.SelectionRange.End;
            fecha_inicio = new DateTime(fecha_fin.Year, fecha_fin.Month, 1);
            //Sacamos la fecha en juliana correspondiente
            int _FechaFinalEnJuliana = ConvertToJulian(fecha_fin);
            int _FechaInicialEnJuliana = ConvertToJulian(fecha_inicio);

            //Sacamos la fecha en juliana correspondiente al año anterior 
            int _FechaFinalEnJulianaAñoAnt = ConvertToJulian(fecha_fin.AddYears(-1));
            int _FechaInicialEnJulianaAñoAnt = ConvertToJulian(fecha_inicio.AddYears(-1));

            //DataTable _DtSR = new DataTable();
            ColumnasSR(out _DtSR);
            for (int i = 0; i < 35; i++)
            {
                DataRow _Rows = _DtSR.NewRow();
                _DtSR.Rows.Add(_Rows);
            }
            _DtSR.Rows[31][0] = "TOTAL";
            _DtSR.Rows[32][0] = "AÑO ANT";
            _DtSR.Rows[33][0] = "DIF %";
            _DtSR.Rows[34][0] = "DIF #";
            DataTable _dtSR;

            //Fechas para pruebas 
            //_FechaInicialEnJuliana = 44256;
            //_FechaFinalEnJuliana = _FechaFinalEnJuliana - 1;//44286; 44256
            //_FechaInicialEnJulianaAñoAnt = 43891;
            //_FechaFinalEnJulianaAñoAnt = _FechaFinalEnJulianaAñoAnt - 1; ; //43921;438

            // Se hace la consulta para sacar el total de mastitis 
            string query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
                         + " FROM DSALUD "
                         + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
                         + " AND vacvaq = 1  AND enfermedad = 1"
                         + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dtSR);
            SacarEspaciosVacios(_dtSR);
            //La primera vez que se manda a llamar LlenarEspaciosVacios se manda un true con el objetivo de
            //sacar las primeras dos columnas incluida la de la fecha, despues es un flase para solo llenar 
            //la columna que nos interesa
            LlenarEspaciosVacios(_DtSR, _dtSR, true, 1);

            //Con este FOR se llena de espacios vacios los días que sean despues de la fecha actual
            for (int i = fecha_fin.Day; i <= 30; i++)
            {
                _DtSR.Rows[i][0] = "NA";
            }

            //Se hace una consulta para saber los metabolicos de la F leche 
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
             + " FROM DSALUD "
             + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
             + " AND vacvaq = 1  AND enfermedad = 5 AND enfdetalle = 1 "
             + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dtSR);
            SacarEspaciosVacios(_dtSR);
            LlenarEspaciosVacios(_DtSR, _dtSR, false, 2);

            //Se hace una consulta para saber los metabolicos de la cetosis
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
             + " FROM DSALUD "
             + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
             + " AND vacvaq = 1  AND enfermedad = 5 AND enfdetalle = 2 "
             + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dtSR);
            SacarEspaciosVacios(_dtSR);
            LlenarEspaciosVacios(_DtSR, _dtSR, false, 3);

            //Se hace una consulta para saber los problemas locomotores, cojas 
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
             + " FROM DSALUD "
             + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
             + " AND vacvaq = 1  AND enfermedad = 4 AND enfdetalle = 1 "
             + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dtSR);
            SacarEspaciosVacios(_dtSR);
            LlenarEspaciosVacios(_DtSR, _dtSR, false, 4);

            //Se hace una consulta para saber los problemas locomotores, laminitis 
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
             + " FROM DSALUD "
             + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
             + " AND vacvaq = 1  AND enfermedad = 4 AND enfdetalle = 2 "
             + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dtSR);
            SacarEspaciosVacios(_dtSR);
            LlenarEspaciosVacios(_DtSR, _dtSR, false, 5);

            //Se hace una consulta para saber los problemas locomotores, gabarro
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
             + " FROM DSALUD "
             + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
             + " AND vacvaq = 1  AND enfermedad = 4 AND enfdetalle = 3 "
             + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dtSR);
            SacarEspaciosVacios(_dtSR);
            LlenarEspaciosVacios(_DtSR, _dtSR, false, 6);

            //Se hace una consulta para saber los problemas digestivos, desplazos 
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
             + " FROM DSALUD "
             + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
             + " AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 1 "
             + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dtSR);
            SacarEspaciosVacios(_dtSR);
            LlenarEspaciosVacios(_DtSR, _dtSR, false, 7);

            //Se hace una consulta para saber los problemas digestivos, empacho
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
             + " FROM DSALUD "
             + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
             + " AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 2 "
             + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dtSR);
            SacarEspaciosVacios(_dtSR);
            LlenarEspaciosVacios(_DtSR, _dtSR, false, 8);

            //Se hace una consulta para saber los problemas digestivos, diarrea
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
             + " FROM DSALUD "
             + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
             + " AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 3 "
             + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dtSR);
            SacarEspaciosVacios(_dtSR);
            LlenarEspaciosVacios(_DtSR, _dtSR, false, 9);

            //Se hace una consulta para saber los problemas reproductivos, retención
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
             + " FROM DSALUD "
             + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
             + " AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 1 "
             + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dtSR);
            SacarEspaciosVacios(_dtSR);
            LlenarEspaciosVacios(_DtSR, _dtSR, false, 10);

            //Se hace una consulta para saber los problemas reproductivos, Metritis
            query = "SELECT  DAY(CDATE(FECHA)), COUNT(*)  "
             + " FROM DSALUD "
             + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
             + " AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 2 "
             + " GROUP BY FECHA  ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dtSR);
            SacarEspaciosVacios(_dtSR);
            LlenarEspaciosVacios(_DtSR, _dtSR, false, 11);

            //Se hace una consulta para saber el numero de vacas preñadas
            query = "SELECT DAY(CDATE(FECHA)),  preñvc  "
             + " FROM VALINVENTARIO "
             + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
             + " ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dtSR);
            SacarEspaciosVacios(_dtSR);
            LlenarEspaciosVacios(_DtSR, _dtSR, false, 13);

            //Se hace una consulta para saber el numero de vacas vacias
            query = "SELECT DAY(CDATE(FECHA)),  vaciavc  "
             + " FROM VALINVENTARIO "
             + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
             + " ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dtSR);
            SacarEspaciosVacios(_dtSR);
            LlenarEspaciosVacios(_DtSR, _dtSR, false, 15);

            //Se hace una consulta para saber el numero de vacas preñadas
            query = "SELECT DAY(CDATE(FECHA)),  preñvq  "
             + " FROM VALINVENTARIO "
             + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
             + " ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dtSR);
            SacarEspaciosVacios(_dtSR);
            LlenarEspaciosVacios(_DtSR, _dtSR, false, 18);

            //Se hace una consulta para saber el numero de vacas vacias
            query = "SELECT DAY(CDATE(FECHA)),  vaciavq  "
             + " FROM VALINVENTARIO "
             + " WHERE FECHA BETWEEN " + _FechaInicialEnJuliana + " AND " + _FechaFinalEnJuliana
             + " ORDER BY FECHA";
            conn.QueryMovGanado(query, out _dtSR);
            SacarEspaciosVacios(_dtSR);
            LlenarEspaciosVacios(_DtSR, _dtSR, false, 20);

            //En este for validamos que si hay ceros  rellene la casilla con vacío y si es nulo que no marque error
            for (int i = 0; i < _DtSR.Rows.Count; i++)
            {
                //Validamos que no sean espacios vacios 
                try
                {
                    int pren = _DtSR.Rows[i][13] != DBNull.Value ? Convert.ToInt32(_DtSR.Rows[i][13]) : 0;
                    int vacias = _DtSR.Rows[i][15] != DBNull.Value ? Convert.ToInt32(_DtSR.Rows[i][15]) : 0;
                    int diag = pren + vacias;
                    decimal porcPren = diag > 0 ? Math.Round(Convert.ToDecimal(pren) / Convert.ToDecimal(diag) * 100) : 0;
                    decimal porcVacias = diag > 0 ? Convert.ToDecimal(vacias) / Convert.ToDecimal(diag) * 100 : 0;
                    string s_pren = pren == 0 ? "" : pren.ToString();
                    string s_vacias = vacias == 0 ? "" : vacias.ToString();
                    string s_diag = diag == 0 ? "" : diag.ToString();
                    string s_porcPren = porcPren == 0 ? "" : porcPren.ToString();
                    string s_porcVacias = porcVacias == 0 ? "" : porcVacias.ToString();

                    _DtSR.Rows[i][13] = s_pren;
                    _DtSR.Rows[i][15] = s_vacias;

                    if (s_diag != "")
                        _DtSR.Rows[i][12] = diag;
                    else
                        _DtSR.Rows[i][12] = DBNull.Value;

                    if (porcPren != 0)
                        _DtSR.Rows[i][14] = porcPren;
                    else
                        _DtSR.Rows[i][14] = DBNull.Value;

                    if (porcVacias != 0)
                        _DtSR.Rows[i][16] = porcVacias;
                    else
                        _DtSR.Rows[i][16] = DBNull.Value;
                }
                catch (Exception ex)
                {

                    Console.WriteLine(ex.Message);

                }

                /* Forma Cañez
                if (_DtSR.Rows[i][13] == DBNull.Value) { }
                else
                {
                    // si son ceros los ponemos como espacios vacios para que no se vea mal en el reporte 
                    if (Convert.ToInt32(_DtSR.Rows[i][13]) == 0)
                    {
                        _DtSR.Rows[i][13] = ""; //_DtSR.Rows[i][15] = "";
                    }
                    else
                    {
                        //En caso de que los valores sean correctos hacemos las operaciones ocrrespondientes 
                        _DtSR.Rows[i][12] = Convert.ToInt32(_DtSR.Rows[i][13]) + Convert.ToInt32(_DtSR.Rows[i][15]);
                        //_DtSR.Rows[i][14] = Decimal.Round((Convert.ToDecimal(_DtSR.Rows[i][13]) / Convert.ToDecimal(_DtSR.Rows[i][12]) * 100));
                        //_DtSR.Rows[i][16] = Decimal.Round((Convert.ToDecimal(_DtSR.Rows[i][15]) / Convert.ToDecimal(_DtSR.Rows[i][12]) * 100));
                        _DtSR.Rows[i][14] = Math.Round((Convert.ToDecimal(_DtSR.Rows[i][13]) / Convert.ToDecimal(_DtSR.Rows[i][12]) * 100));
                        _DtSR.Rows[i][16] = Math.Round((Convert.ToDecimal(_DtSR.Rows[i][15]) / Convert.ToDecimal(_DtSR.Rows[i][12]) * 100));
                    }

                }
                */

                try
                {
                    int pren = _DtSR.Rows[i][18] != DBNull.Value ? Convert.ToInt32(_DtSR.Rows[i][18]) : 0;
                    int vacias = _DtSR.Rows[i][20] != DBNull.Value ? Convert.ToInt32(_DtSR.Rows[i][20]) : 0;
                    int diag = pren + vacias;
                    decimal porcPren = diag > 0 ? Math.Round(Convert.ToDecimal(pren) / Convert.ToDecimal(diag) * 100) : 0;
                    decimal porcVacias = diag > 0 ? Convert.ToDecimal(vacias) / Convert.ToDecimal(diag) * 100 : 0;
                    string s_pren = pren == 0 ? "" : pren.ToString();
                    string s_vacias = vacias == 0 ? "" : vacias.ToString();
                    string s_diag = diag == 0 ? "" : diag.ToString();
                    string s_porcPren = porcPren == 0 ? "" : porcPren.ToString();
                    string s_porcVacias = porcVacias == 0 ? "" : porcVacias.ToString();

                    //_DtSR.Rows[i][17] = s_diag;
                    _DtSR.Rows[i][18] = s_pren;
                    _DtSR.Rows[i][20] = s_vacias;

                    if (s_diag != "")
                        _DtSR.Rows[i][17] = diag;
                    else
                        _DtSR.Rows[i][17] = DBNull.Value;

                    if (porcPren != 0)
                        _DtSR.Rows[i][19] = porcPren;
                    else
                        _DtSR.Rows[i][19] = DBNull.Value;

                    if (porcVacias != 0)
                        _DtSR.Rows[i][21] = porcVacias;
                    else
                        _DtSR.Rows[i][21] = DBNull.Value;
                }
                catch (Exception ex)
                {

                    Console.WriteLine(ex.Message);

                }
                /*Forma cañez
                if (_DtSR.Rows[i][18] == DBNull.Value) { }
                else
                {
                    if (Convert.ToInt32(_DtSR.Rows[i][18]) == 0)
                    {
                        _DtSR.Rows[i][18] = ""; _DtSR.Rows[i][20] = "";
                    }
                    else
                    {
                        _DtSR.Rows[i][17] = Convert.ToInt32(_DtSR.Rows[i][18]) + Convert.ToInt32(_DtSR.Rows[i][20]);
                        //_DtSR.Rows[i][19] = Decimal.Round((Convert.ToDecimal(_DtSR.Rows[i][18]) / Convert.ToDecimal(_DtSR.Rows[i][17]) * 100));
                        //_DtSR.Rows[i][21] = Decimal.Round((Convert.ToDecimal(_DtSR.Rows[i][20]) / Convert.ToDecimal(_DtSR.Rows[i][17]) * 100));
                        _DtSR.Rows[i][19] = Math.Round((Convert.ToDecimal(_DtSR.Rows[i][18]) / Convert.ToDecimal(_DtSR.Rows[i][17]) * 100));
                        _DtSR.Rows[i][21] = Math.Round((Convert.ToDecimal(_DtSR.Rows[i][20]) / Convert.ToDecimal(_DtSR.Rows[i][17]) * 100));
                    }
                }
                */
            }

            //Sacamos al renglon AñoAnt haciendo las mismas consultas pero con respecto al año anterior 
            //Mastitis del año anterior 
            query = "SELECT SUM(t.NUMVACAS) FROM (SELECT DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS"
               + " FROM DSALUD WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
               + " AND vacvaq = 1  AND enfermedad = 1 "
               + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dtSR);
            _DtSR.Rows[32][1] = _dtSR.Rows[0][0];

            //Metabólicos F_Leche del año anterior 
            query = "SELECT SUM(t.NUMVACAS) FROM (SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS"
               + " FROM DSALUD "
               + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
               + " AND vacvaq = 1  AND enfermedad = 5 AND enfdetalle = 1 "
               + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dtSR);
            _DtSR.Rows[32][2] = _dtSR.Rows[0][0];

            //Se hace una consulta para saber los metabolicos de la cetosis del año anterior 
            query = "SELECT SUM(t.NUMVACAS) FROM(SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS"
             + " FROM DSALUD "
             + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
             + " AND vacvaq = 1  AND enfermedad = 5 AND enfdetalle = 2 "
             + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dtSR);
            _DtSR.Rows[32][3] = _dtSR.Rows[0][0];

            //Se hace una consulta para saber los problemas locomotores, cojas del año anterior 
            query = "SELECT SUM(t.NUMVACAS) FROM(SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS"
            + " FROM DSALUD "
            + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
            + " AND vacvaq = 1  AND enfermedad = 4 AND enfdetalle = 1 "
            + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dtSR);
            _DtSR.Rows[32][4] = _dtSR.Rows[0][0];

            //Se hace una consulta para saber los problemas locomotores, laminitis año anterior 
            query = "SELECT SUM(t.NUMVACAS) FROM(SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS"
            + " FROM DSALUD "
            + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
            + " AND vacvaq = 1  AND enfermedad = 4 AND enfdetalle = 2 "
            + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dtSR);
            _DtSR.Rows[32][5] = _dtSR.Rows[0][0];

            //Se hace una consulta para saber los problemas locomotores, gabarro año anterior 
            query = "SELECT SUM(t.NUMVACAS) FROM(SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS"
            + " FROM DSALUD "
            + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
            + " AND vacvaq = 1  AND enfermedad = 4 AND enfdetalle = 3 "
            + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dtSR);
            _DtSR.Rows[32][6] = _dtSR.Rows[0][0];

            //Se hace una consulta para saber los problemas digestivos, desplazos año anterior 
            query = "SELECT SUM(t.NUMVACAS) FROM(SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS"
            + " FROM DSALUD "
            + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
            + " AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 1 "
            + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dtSR);
            _DtSR.Rows[32][7] = _dtSR.Rows[0][0];

            //Se hace una consulta para saber los problemas digestivos, empacho año anterior 
            query = "SELECT SUM(t.NUMVACAS) FROM(SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS"
            + " FROM DSALUD "
            + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
            + " AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 2 "
            + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dtSR);
            _DtSR.Rows[32][8] = _dtSR.Rows[0][0];

            //Se hace una consulta para saber los problemas digestivos, diarrea año anterior 
            query = "SELECT SUM(t.NUMVACAS) FROM(SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS"
            + " FROM DSALUD "
            + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
            + " AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 3 "
            + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dtSR);
            _DtSR.Rows[32][9] = _dtSR.Rows[0][0];

            //Se hace una consulta para saber los problemas reproductivos, retención año anterior 
            query = "SELECT SUM(t.NUMVACAS) FROM(SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS"
            + " FROM DSALUD "
            + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
            + " AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 1 "
            + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dtSR);
            _DtSR.Rows[32][10] = _dtSR.Rows[0][0];

            //Se hace una consulta para saber los problemas reproductivos, metritis año anterior 
            query = "SELECT SUM(t.NUMVACAS) FROM(SELECT  DAY(CDATE(FECHA)), COUNT(*) AS NUMVACAS"
            + " FROM DSALUD "
            + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
            + " AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 2 "
            + " GROUP BY FECHA  ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dtSR);
            _DtSR.Rows[32][11] = _dtSR.Rows[0][0];

            //Se hace una consulta para saber el numero de vacas preñadas año anterior 
            query = "SELECT SUM(t.NUMVACAS) FROM(SELECT  DAY(CDATE(FECHA)), preñvc AS NUMVACAS"
            + " FROM VALINVENTARIO "
            + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
            + " ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dtSR);
            if (Convert.ToString(_dtSR.Rows[0][0]) == "0") { _DtSR.Rows[32][13] = ""; } else { _DtSR.Rows[32][13] = _dtSR.Rows[0][0]; }

            //Se hace una consulta para saber el numero de vacas vacias año anterior 
            query = "SELECT SUM(t.NUMVACAS) FROM(SELECT  DAY(CDATE(FECHA)), vaciavc AS NUMVACAS"
            + " FROM VALINVENTARIO "
            + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
            + " ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dtSR);
            if (Convert.ToString(_dtSR.Rows[0][0]) == "0") { _DtSR.Rows[32][15] = ""; } else { _DtSR.Rows[32][15] = _dtSR.Rows[0][0]; }

            //Se hace una consulta para saber el numero de vaquilas preñadas año anterior 
            query = "SELECT SUM(t.NUMVACAS) FROM(SELECT  DAY(CDATE(FECHA)), preñvq AS NUMVACAS"
            + " FROM VALINVENTARIO "
            + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
            + " ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dtSR);
            if (Convert.ToString(_dtSR.Rows[0][0]) == "0") { _DtSR.Rows[32][18] = ""; } else { _DtSR.Rows[32][18] = _dtSR.Rows[0][0]; }


            //Se hace una consulta para saber el numero de vaquillas vacias año anterior 
            query = "SELECT SUM(t.NUMVACAS) FROM(SELECT  DAY(CDATE(FECHA)), vaciavq  AS NUMVACAS"
            + " FROM VALINVENTARIO "
            + " WHERE FECHA BETWEEN " + _FechaInicialEnJulianaAñoAnt + " AND " + _FechaFinalEnJulianaAñoAnt
            + " ORDER BY FECHA) t";
            conn.QueryMovGanado(query, out _dtSR);
            if (Convert.ToString(_dtSR.Rows[0][0]) == "0") { _DtSR.Rows[32][20] = ""; } else { _DtSR.Rows[32][20] = _dtSR.Rows[0][0]; }

            //En este for validamos que si hay ceros  rellene la casilla con vacío y si es nulo que no marque error
            if (Convert.ToString(_DtSR.Rows[32][20]) == "" && Convert.ToString(_DtSR.Rows[32][18]) == "") { }
            else
            {
                //En caso de que alguno de los dos campos sea cero pero el otro tenga valor se identifica el que tiene "", se le regresa el valor cero y después 
                //se vuelve a poner como "" para que no afecte el reporte, el primer if es para las vacas y el segundo if para las vaquillas 
                if (Convert.ToString(_DtSR.Rows[32][18]) == "") { _DtSR.Rows[32][18] = "0"; }
                if (Convert.ToString(_DtSR.Rows[32][20]) == "") { _DtSR.Rows[32][20] = "0"; }
                _DtSR.Rows[32][17] = Convert.ToInt32(_DtSR.Rows[32][18]) + Convert.ToInt32(_DtSR.Rows[32][20]);
                _DtSR.Rows[32][19] = Math.Round((Convert.ToDecimal(_DtSR.Rows[32][18]) / Convert.ToDecimal(_DtSR.Rows[32][17]) * 100));
                _DtSR.Rows[32][21] = Math.Round((Convert.ToDecimal(_DtSR.Rows[32][20]) / Convert.ToDecimal(_DtSR.Rows[32][17]) * 100));
                if (Convert.ToString(_DtSR.Rows[32][20]) == "0") { _DtSR.Rows[32][20] = ""; }
                if (Convert.ToString(_DtSR.Rows[32][18]) == "0") { _DtSR.Rows[32][18] = ""; }
            }
            if (Convert.ToString(_DtSR.Rows[32][15]) == "" && Convert.ToString(_DtSR.Rows[32][13]) == "") { }
            else
            {
                if (Convert.ToString(_DtSR.Rows[32][15]) == "") { _DtSR.Rows[32][15] = "0"; }
                if (Convert.ToString(_DtSR.Rows[32][13]) == "") { _DtSR.Rows[32][13] = "0"; }
                _DtSR.Rows[32][12] = Convert.ToInt32(_DtSR.Rows[32][13]) + Convert.ToInt32(_DtSR.Rows[32][15]);
                _DtSR.Rows[32][14] = Math.Round((Convert.ToDecimal(_DtSR.Rows[32][13]) / Convert.ToDecimal(_DtSR.Rows[32][12]) * 100));
                _DtSR.Rows[32][16] = Math.Round((Convert.ToDecimal(_DtSR.Rows[32][15]) / Convert.ToDecimal(_DtSR.Rows[32][12]) * 100));
                if (Convert.ToString(_DtSR.Rows[32][15]) == "0") { _DtSR.Rows[32][15] = ""; }
                if (Convert.ToString(_DtSR.Rows[32][13]) == "0") { _DtSR.Rows[32][13] = ""; }
            }

            //Este for nos permite evaluar los valores que hay en Total y Total año anterior, si uno de los dos no tiene valor no hace nada
            // pero si ambos tienen valor entonces pasa a hacer las operaciones.
            for (int i = 0; i <= 20; i++)
            {
                if (Convert.ToString(_DtSR.Rows[31][i + 1]) == "" || Convert.ToString(_DtSR.Rows[32][i + 1]) == "" || Convert.ToString(_DtSR.Rows[31][i + 1]) == "0" || Convert.ToString(_DtSR.Rows[32][i + 1]) == "0") { }
                else
                {
                    //_DtSR.Rows[33][i + 1] = Decimal.Round(((Convert.ToDecimal(_DtSR.Rows[32][i + 1]) / Convert.ToDecimal(_DtSR.Rows[31][i + 1])) - 1) * -100);
                    //_DtSR.Rows[33][i + 1] = Math.Round(((Convert.ToDecimal(_DtSR.Rows[32][i + 1]) / Convert.ToDecimal(_DtSR.Rows[31][i + 1])) - 1) * -100);
                    _DtSR.Rows[34][i + 1] = Convert.ToInt32(_DtSR.Rows[31][i + 1]) - Convert.ToInt32(_DtSR.Rows[32][i + 1]);
                    _DtSR.Rows[33][i + 1] = Math.Round(((Convert.ToDecimal(_DtSR.Rows[34][i + 1]) / Convert.ToDecimal(_DtSR.Rows[32][i + 1]))) * 100);
                    if (Convert.ToDouble(_DtSR.Rows[34][i + 1]) > 0)
                    {
                        _DtSR.Rows[33][i + 1] = Math.Abs(Convert.ToDouble(_DtSR.Rows[33][i + 1]));
                    }
                }
            }
        }
        private void ColumnasSR(out DataTable dt)
        {
            dt = new DataTable();
            dt.Columns.Add("DIA").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("MASTITIS").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("M_FLECHE").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("M_CETOSIS").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("L_COJAS").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("L_LAMINITITIS").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("L_GABARRO").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("D_DESPLAZOS").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("D_EMPACHO").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("D_DIARREA").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("R_RETENCION").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("R_METRITIS").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VA_DIAG").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VA_PREN").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("VA_PPOR").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VA_VACIAS").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("VA_VPOR").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VAQ_PART").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VAQ_PRE").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("VAQ_PPOR").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("VAQ_VACIAS").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("VAQ_VPOR").DataType = System.Type.GetType("System.Int32");


        }

        private void PonderadosHoja1(DataTable dtHoja1, DataTable DtPrecioLeche)
        {
            // arreglo con valores para las ponderaciones 
            //0, 1, 2, 3, 4, 5, 6, 7, 8, 9,10,11,12,13,14,15,16,17,18,19,20,21,22
            double[] Acumulador = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            /*
            Posicion 0 es la sumatortia del porcentaje de lactancia entre 100 multiplicado por el ordeño 
            posicion 1 es la sumatortia del porcentaje de la proteina entre 100 multiplicado por el ordeño 
            posicion 2 es la sumatortia de UERA entre 100 multiplicado por el ordeño
            posicion 3 es la sumatortia de la grasa entre 100 multiplicado por el ordeño
            posicion 4 es la sumatortia del CCS entre 100 multiplicado por el ordeño
            posicion 5 es la sumatortia del CDT entre 100 multiplicado por el ordeño
            posicion 5 es la sumatortia del CDT entre 100 multiplicado por el ordeño
            posicion 6 es la sumatoria del total entre el ordeño todo multiplicado por el precio de la leche 
            posicion 7 es un contador para saber si hay valores en el día actual 
            posicion 8 es una sumatoria de la porducción 
            posicion 9 es una sumatoria del valor X
            posicion 10 es una sumatoria del MS
            posicion 11 es una sumatoria del valor X multiplicado por el precio de la leche, como el valor de la leche puede cambiar no se usa la posicion 9 
            posicion 12 es un contador para saber si hay valores en el día actual con respecto al valor X
            posicion 14 es un acumulador del total (primera columna) que se suma siempre y cuango haya valor en la grasa.
            posicion 15 es un acumulador del total (primera columna) que se suma siempre y cuango haya valor en urea.
            posicion 16 es un acumulador del total (primera columna) que se suma siempre y cuango haya valor en ccs.
            posicion 17 es un acumulador del total (primera columna) que se suma siempre y cuango haya valor en CDT.
            posicion 18 es un acumulador del total (primera columna) que se suma siempre y cuango haya valor en %Lact
            posicion 19 es un acumulador del total (primera columna) que se suma siempre y cuango haya valor en %PROT
            Posicion 20 es la sumatoria del totoal
            Posicion 21 es la sumatoria del ordeño 
            Posicion 22 es la multiplicación del ordeño por el precio de procucción 
            */
            for (int i = 0; i <= 30; i++)
            {
                if (dtHoja1.Rows[i][7] != DBNull.Value)
                {
                    if (Convert.ToDouble(dtHoja1.Rows[i][7]) > 0)
                    {
                        Acumulador[15] += dtHoja1.Rows[i][1] != DBNull.Value ? (Convert.ToDouble(dtHoja1.Rows[i][1])) : 0;
                    }
                }
                if (dtHoja1.Rows[i][8] != DBNull.Value)
                {
                    if (Convert.ToDouble(dtHoja1.Rows[i][8]) > 0)
                    {
                        Acumulador[14] += dtHoja1.Rows[i][1] != DBNull.Value ? (Convert.ToDouble(dtHoja1.Rows[i][1])) : 0;
                    }

                }
                if (dtHoja1.Rows[i][9] != DBNull.Value)
                {
                    if (Convert.ToDouble(dtHoja1.Rows[i][9]) > 0)
                    {
                        Acumulador[16] += dtHoja1.Rows[i][1] != DBNull.Value ? (Convert.ToDouble(dtHoja1.Rows[i][1])) : 0;
                    }
                }
                if (dtHoja1.Rows[i][10] != DBNull.Value)
                {
                    if (Convert.ToDouble(dtHoja1.Rows[i][10]) > 0)
                    {
                        Acumulador[17] += dtHoja1.Rows[i][1] != DBNull.Value ? (Convert.ToDouble(dtHoja1.Rows[i][1])) : 0;
                    }
                }
                if (dtHoja1.Rows[i][5] != DBNull.Value)
                {
                    if (Convert.ToDouble(dtHoja1.Rows[i][5]) > 0)
                    {
                        Acumulador[18] += dtHoja1.Rows[i][1] != DBNull.Value ? (Convert.ToDouble(dtHoja1.Rows[i][1])) : 0;
                    }
                }
                if (dtHoja1.Rows[i][6] != DBNull.Value)
                {
                    if (Convert.ToDouble(dtHoja1.Rows[i][6]) > 0)
                    {
                        Acumulador[19] += dtHoja1.Rows[i][1] != DBNull.Value ? (Convert.ToDouble(dtHoja1.Rows[i][1])) : 0;
                    }
                }
            }
            for (int i = 1; i <= 34; i++)
            {
                Acumulador[13] = 0;
                Acumulador[20] = 0;
                Acumulador[21] = 0;
                Acumulador[22] = 0;
                int _seguro = 0, _contador = 0; //El seguro nos ayuda a identificar si hay un valor, cómo la celda donde esta el resultado tiene un valor nulo
                //es necesario guardar un número cero para poder hacer la suma, en caso de no contener valores somplemene se queda vacío y no con un cero.
                for (int j = 0; j <= 30; j++)
                {
                    if (dtHoja1.Rows[j][i] == DBNull.Value) { }
                    else
                    {
                        //Si hay valores negativos o cero, validamos para que no los mueste en el reporte ni los cuente para la sumatoria.
                        if (i >= 18 && i <= 22)
                        {
                            if (_seguro == 0)
                            {
                                dtHoja1.Rows[32][i] = 0;
                                _seguro = 1;
                            }
                            dtHoja1.Rows[32][i] = Convert.ToDouble(dtHoja1.Rows[32][i]) + Convert.ToDouble(dtHoja1.Rows[j][i]);
                            _contador++;//Cada vez que hay un valor no nulo cuenta 
                            Acumulador[13] += dtHoja1.Rows[j][1] != DBNull.Value ? Convert.ToDouble(Convert.ToDouble(dtHoja1.Rows[j][i])) * Convert.ToDouble(Convert.ToDouble(dtHoja1.Rows[j][1])) : 0;
                            Acumulador[20] += dtHoja1.Rows[j][1] != DBNull.Value ? Convert.ToDouble(Convert.ToDouble(dtHoja1.Rows[j][1])) : 0;
                        }
                        else
                        {
                            if (Convert.ToDouble(dtHoja1.Rows[j][i]) <= 0)
                            {
                                dtHoja1.Rows[j][i] = DBNull.Value;
                            }
                            else
                            {
                                if (_seguro == 0)
                                {
                                    dtHoja1.Rows[32][i] = 0;
                                    _seguro = 1;
                                }
                                dtHoja1.Rows[32][i] = Convert.ToDouble(dtHoja1.Rows[32][i]) + Convert.ToDouble(dtHoja1.Rows[j][i]);
                                _contador++;//Cada vez que hay un valor no nulo cuenta 
                                Acumulador[13] += dtHoja1.Rows[j][1] != DBNull.Value ? Convert.ToDouble(Convert.ToDouble(dtHoja1.Rows[j][i])) * Convert.ToDouble(Convert.ToDouble(dtHoja1.Rows[j][1])) : 0;
                                Acumulador[20] += dtHoja1.Rows[j][1] != DBNull.Value ? Convert.ToDouble(Convert.ToDouble(dtHoja1.Rows[j][1])) : 0;
                                Acumulador[21] += dtHoja1.Rows[j][2] != DBNull.Value ? Convert.ToDouble(Convert.ToDouble(dtHoja1.Rows[j][2])) : 0;
                                Acumulador[22] += dtHoja1.Rows[j][2] != DBNull.Value ? Convert.ToDouble(Convert.ToDouble(dtHoja1.Rows[j][i])) * Convert.ToDouble(Convert.ToDouble(dtHoja1.Rows[j][2])) : 0;
                            }
                        }
                    }
                }
                if (dtHoja1.Rows[32][i] == DBNull.Value)//Validamos que no haya un valor nulo, en caso de ser nulo no hace nada 
                {

                }
                else//De otro modo hace la operación para sacar el promedio
                {
                    if (i == 26)
                    {
                        Acumulador[10] = Convert.ToInt32(dtHoja1.Rows[32][i]);
                    }
                    if (i == 13)
                    {
                        Acumulador[9] = Convert.ToInt32(dtHoja1.Rows[32][i]);
                    }
                    if (i == 31)
                    {
                        Acumulador[8] = Convert.ToInt32(dtHoja1.Rows[32][i]);
                    }
                    if (_contador == 0)
                    {
                        dtHoja1.Rows[32][i] = 0;
                    }
                    else
                    {
                        dtHoja1.Rows[32][i] = Convert.ToDouble(dtHoja1.Rows[32][i]) / _contador;
                    }
                    if (i >= 24 && i <= 32)
                    {
                        if (dtHoja1.Rows[31][1] != DBNull.Value)
                        {
                            // dtHoja1.Rows[32][i] = Acumulador[13] / Convert.ToDouble(dtHoja1.Rows[31][1]);
                            dtHoja1.Rows[32][i] = Acumulador[13] / Acumulador[20];
                        }
                        else
                        {
                            dtHoja1.Rows[32][i] = DBNull.Value;
                        }
                        if (i == 31) { dtHoja1.Rows[32][i] = Acumulador[21] != 0 ? Acumulador[22] / Acumulador[21] : 0; }
                        //Se toma en cuenta el 27 (S/A) porque debe de tomar en cuenta la cantidad de días independientemente de si hay o no valores
                        if (i == 27)
                        {
                            double _SA = 0;
                            for (int j = 0; j < 31; j++)
                            {
                                _SA += dtHoja1.Rows[j][i] != DBNull.Value ? Convert.ToDouble(dtHoja1.Rows[j][i]) : 0;
                            }
                            _SA /= monthCalendar1.SelectionRange.Start.Day;
                            dtHoja1.Rows[32][i] = _SA != 0 ? _SA.ToString() : 0.ToString();
                        }
                    }
                }
            }
            // una vez sacado los promedios susituimos los valores en los promedio que no estan ponderado 
            for (int i = 0; i < 31; i++)
            {
                // en este ciclo de días revisamos si hay valores en las celdas que nos importan, de ser así tomamos los datos y se hace el ponderado 
                if (dtHoja1.Rows[i][5] == DBNull.Value) { }
                else
                {
                    Acumulador[0] = dtHoja1.Rows[i][1] != DBNull.Value ? ((Convert.ToDouble(dtHoja1.Rows[i][5]) / 100) * Convert.ToDouble(dtHoja1.Rows[i][1])) + Acumulador[0] : 0;
                }
                if (dtHoja1.Rows[i][6] == DBNull.Value) { }
                else
                {
                    Acumulador[1] = dtHoja1.Rows[i][1] != DBNull.Value ? ((Convert.ToDouble(dtHoja1.Rows[i][6]) / 100) * Convert.ToDouble(dtHoja1.Rows[i][1])) + Acumulador[1] : 0;
                }
                if (dtHoja1.Rows[i][7] == DBNull.Value) { }
                else
                {
                    Acumulador[2] = dtHoja1.Rows[i][1] != DBNull.Value ? ((Convert.ToDouble(dtHoja1.Rows[i][7]) / 100) * Convert.ToDouble(dtHoja1.Rows[i][1])) + Acumulador[2] : 0;
                }
                if (dtHoja1.Rows[i][8] == DBNull.Value) { }
                else
                {
                    Acumulador[3] = dtHoja1.Rows[i][1] != DBNull.Value ? ((Convert.ToDouble(dtHoja1.Rows[i][8]) / 100) * Convert.ToDouble(dtHoja1.Rows[i][1])) + Acumulador[3] : 0;
                }
                if (dtHoja1.Rows[i][9] == DBNull.Value) { }
                else
                {
                    Acumulador[4] = dtHoja1.Rows[i][1] != DBNull.Value ? ((Convert.ToDouble(dtHoja1.Rows[i][9]) / 100) * Convert.ToDouble(dtHoja1.Rows[i][1])) + Acumulador[4] : 0;
                }
                if (dtHoja1.Rows[i][10] == DBNull.Value) { }
                else
                {
                    Acumulador[5] = dtHoja1.Rows[i][1] != DBNull.Value ? ((Convert.ToDouble(dtHoja1.Rows[i][10]) / 100) * Convert.ToDouble(dtHoja1.Rows[i][1])) + Acumulador[5] : 0;
                }
                if (dtHoja1.Rows[i][1] == DBNull.Value || dtHoja1.Rows[i][2] == DBNull.Value || dtHoja1.Rows[i][31] == DBNull.Value) { }
                else
                {
                    if (DtPrecioLeche.Rows[i][1] != DBNull.Value)
                    {
                        Acumulador[6] = Acumulador[6] + ((Convert.ToDouble(dtHoja1.Rows[i][1]) / (Convert.ToDouble(dtHoja1.Rows[i][2]))) * Convert.ToDouble(DtPrecioLeche.Rows[i][1]));
                        Acumulador[7]++;
                    }
                }
                if (dtHoja1.Rows[i][13] == DBNull.Value || dtHoja1.Rows[i][1] == DBNull.Value || dtHoja1.Rows[i][31] == DBNull.Value) { }
                else
                {
                    if (DtPrecioLeche.Rows[i][1] != DBNull.Value)
                    {
                        Acumulador[11] = (Convert.ToDouble(dtHoja1.Rows[i][13]) * Convert.ToDouble(DtPrecioLeche.Rows[i][1])) + Acumulador[11];
                        Acumulador[12]++;
                    }
                }
            }
            //Se agregan los nuevos valores ya ponderizados de las tablas.
            if (Acumulador[8] != 0)
            {
                dtHoja1.Rows[32][22] = (Acumulador[9] - Acumulador[8]) / Acumulador[8];
                dtHoja1.Rows[32][21] = (Acumulador[11] / Acumulador[8]);
                dtHoja1.Rows[32][18] = (Acumulador[6] / Acumulador[8]);
            }
            if (Acumulador[12] != 0) { dtHoja1.Rows[32][19] = (Acumulador[11] - Acumulador[8]) / Acumulador[12]; }
            if (Acumulador[10] != 0) { dtHoja1.Rows[32][20] = (Acumulador[9] / Acumulador[10]); }
            if (Acumulador[7] != 0) { dtHoja1.Rows[32][22] = (Acumulador[6] - Acumulador[8]) / Acumulador[7]; }

            dtHoja1.Rows[32][5] = Acumulador[18] != 0 ? (Acumulador[0] / Acumulador[18]) * 100 : 0;
            dtHoja1.Rows[32][6] = Acumulador[19] != 0 ? (Acumulador[1] / Acumulador[19]) * 100 : 0;
            dtHoja1.Rows[32][7] = Acumulador[15] != 0 ? (Acumulador[2] / Acumulador[15]) * 100 : 0;
            dtHoja1.Rows[32][8] = Acumulador[14] != 0 ? (Acumulador[3] / Acumulador[14]) * 100 : 0;
            dtHoja1.Rows[32][9] = Acumulador[16] != 0 ? (Acumulador[4] / Acumulador[16]) * 100 : 0;
            dtHoja1.Rows[32][10] = Acumulador[17] != 0 ? (Acumulador[5] / Acumulador[17]) * 100 : 0;

            if (dtHoja1.Rows[32][2] != DBNull.Value)
            {
                dtHoja1.Rows[32][13] = dtHoja1.Rows[32][14] != DBNull.Value ? (Convert.ToDouble(dtHoja1.Rows[32][14])) / (Convert.ToDouble(dtHoja1.Rows[32][2])) : 0;
            }
        }

        private void ValoresReporteDiarioProduccion(DateTime Finicio, DateTime FFinal, DataTable dt1)
        {
            bendiciones.ReportePeriodo sobrante = new bendiciones.ReportePeriodo();
            int hcorte = 0, horas;
            Hora_Corte(out horas, out hcorte);
            List<bendiciones.ReportePeriodo> listaIngredientes = new List<bendiciones.ReportePeriodo>();
            List<bendiciones.ReportePeriodo> listaForraje = new List<bendiciones.ReportePeriodo>();
            List<bendiciones.ReportePeriodo> listaAlimento = new List<bendiciones.ReportePeriodo>();
            List<bendiciones.ReportePeriodo> listaAgua = new List<bendiciones.ReportePeriodo>();
            bendiciones.IndicadorReportePeriodo indicadores = new bendiciones.IndicadorReportePeriodo();
            GTH.ReportePeriodo(ran_id.ToString(), horas, "10,11,12,13", Finicio.Date, FFinal.Date, out listaIngredientes, out indicadores, out sobrante);

            if (FFinal.Year == fechaFinDTP.Year)
                prom_ordeño = indicadores;

            try
            {
                dt1.Rows[32]["X"] = indicadores.MEDIA;
                dt1.Rows[32]["ILCAVTA"] = indicadores.ILCA_VENTA;
                dt1.Rows[32]["ICVTA"] = indicadores.IC_VENTA;
                dt1.Rows[32]["EA"] = indicadores.EA;
                dt1.Rows[32]["ILCAPROD"] = indicadores.ILCA_PRODUCCION;
                dt1.Rows[32]["ICPROD"] = indicadores.IC_PRODUCCION;
                dt1.Rows[32]["PRECIOL"] = indicadores.PRECIOL;
                dt1.Rows[32]["MH"] = indicadores.MH;
                dt1.Rows[32]["PORCMS"] = indicadores.PORCENTAJEMS;
                dt1.Rows[32]["MS"] = indicadores.MS;
                dt1.Rows[32]["SA"] = indicadores.SA;
                dt1.Rows[32]["MSS"] = indicadores.MSS;
                dt1.Rows[32]["EAS"] = indicadores.EAS;
                dt1.Rows[32]["PRECIOPROD"] = indicadores.COSTO;
                dt1.Rows[32]["PRECIOMS"] = indicadores.PRECIOKGMS;
            }
            catch
            {

            }
        }

        private void ValoresReporteDiarioHojaDos(DateTime Finicio, DateTime FFinal, DataTable dt)
        {
            bendiciones.ReportePeriodo sobrante = new bendiciones.ReportePeriodo();
            int hcorte = 0, horas;
            Hora_Corte(out horas, out hcorte);
            List<bendiciones.ReportePeriodo> listaIngredientes = new List<bendiciones.ReportePeriodo>();
            List<bendiciones.ReportePeriodo> listaForraje = new List<bendiciones.ReportePeriodo>();
            List<bendiciones.ReportePeriodo> listaAlimento = new List<bendiciones.ReportePeriodo>();
            List<bendiciones.ReportePeriodo> listaAgua = new List<bendiciones.ReportePeriodo>();
            bendiciones.IndicadorReportePeriodo indicadoresJaulas = new bendiciones.IndicadorReportePeriodo();
            bendiciones.IndicadorReportePeriodo indicadoresReto = new bendiciones.IndicadorReportePeriodo();
            bendiciones.IndicadorReportePeriodo indicadoresDesteteUno = new bendiciones.IndicadorReportePeriodo();
            bendiciones.IndicadorReportePeriodo indicadoresDesteteDos = new bendiciones.IndicadorReportePeriodo();
            bendiciones.IndicadorReportePeriodo indicadoresSecas = new bendiciones.IndicadorReportePeriodo();
            bendiciones.IndicadorReportePeriodo indicadoresVaquillas = new bendiciones.IndicadorReportePeriodo();
            GTH.ReportePeriodo(ran_id.ToString(), horas, "31", Finicio.AddDays(1).Date, FFinal.Date, out listaIngredientes, out indicadoresJaulas, out sobrante);
            GTH.ReportePeriodo(ran_id.ToString(), horas, "22", Finicio.AddDays(1).Date, FFinal.Date, out listaIngredientes, out indicadoresReto, out sobrante);
            GTH.ReportePeriodo(ran_id.ToString(), horas, "32", Finicio.AddDays(1).Date, FFinal.Date, out listaIngredientes, out indicadoresDesteteUno, out sobrante);
            GTH.ReportePeriodo(ran_id.ToString(), horas, "33", Finicio.AddDays(1).Date, FFinal.Date, out listaIngredientes, out indicadoresDesteteDos, out sobrante);
            GTH.ReportePeriodo(ran_id.ToString(), horas, "21", Finicio.AddDays(1).Date, FFinal.Date, out listaIngredientes, out indicadoresSecas, out sobrante);
            GTH.ReportePeriodo(ran_id.ToString(), horas, "34", Finicio.AddDays(1).Date, FFinal.Date, out listaIngredientes, out indicadoresVaquillas, out sobrante);

            //Guardar
            if (FFinal.Year == fechaFinDTP.Year)
            {
                prom_reto = indicadoresReto;
                prom_secas = indicadoresSecas;
                prom_destete1 = indicadoresDesteteUno;
                prom_destete2 = indicadoresDesteteDos;
                prom_vquillas = indicadoresVaquillas;
            }

            //Asignar a variables de promedios
            //Jaulas
            dt.Rows[32]["PRECIOJ"] = indicadoresJaulas.COSTO;
            //Destete uno
            dt.Rows[32]["MH2"] = indicadoresDesteteUno.MH;
            dt.Rows[32]["PRECIO2"] = indicadoresDesteteUno.COSTO;
            dt.Rows[32]["PORCMS2"] = indicadoresDesteteUno.PORCENTAJEMS;
            dt.Rows[32]["MS2"] = indicadoresDesteteUno.MS;
            dt.Rows[32]["PRECIOMS2"] = indicadoresDesteteUno.PRECIOKGMS;
            //Destete Dos
            dt.Rows[32]["MH7"] = indicadoresDesteteDos.MH;
            dt.Rows[32]["PRECIO7"] = indicadoresDesteteDos.COSTO;
            dt.Rows[32]["PORCMS7"] = indicadoresDesteteDos.PORCENTAJEMS;
            dt.Rows[32]["MS7"] = indicadoresDesteteDos.MS;
            dt.Rows[32]["PRECIOMS7"] = indicadoresDesteteDos.PRECIOKGMS;
            //Vaquillas
            dt.Rows[32]["MH13"] = indicadoresVaquillas.MH;
            dt.Rows[32]["PRECIO13"] = indicadoresVaquillas.COSTO;
            dt.Rows[32]["PORCMS13"] = indicadoresVaquillas.PORCENTAJEMS;
            dt.Rows[32]["MS13"] = indicadoresVaquillas.MS;
            dt.Rows[32]["PRECIOMS13"] = indicadoresVaquillas.PRECIOKGMS;
            //Secas
            dt.Rows[32]["MHSECAS"] = indicadoresSecas.MH;
            dt.Rows[32]["PORCMSSECAS"] = indicadoresSecas.PORCENTAJEMS;
            dt.Rows[32]["MSSECAS"] = indicadoresSecas.MS;
            dt.Rows[32]["SASECAS"] = indicadoresSecas.SA;
            dt.Rows[32]["MSSSECAS"] = indicadoresSecas.MSS;
            dt.Rows[32]["PRECIOSECAS"] = indicadoresSecas.COSTO;
            dt.Rows[32]["PRECIOMSSECAS"] = indicadoresSecas.PRECIOKGMS;
            //Reto
            dt.Rows[32]["MHreto"] = indicadoresReto.MH;
            dt.Rows[32]["PORCMSRETO"] = indicadoresReto.PORCENTAJEMS;
            dt.Rows[32]["MSRETO"] = indicadoresReto.MS;
            dt.Rows[32]["SARETO"] = indicadoresReto.SA;
            dt.Rows[32]["MSSRETO"] = indicadoresReto.MSS;
            dt.Rows[32]["PRECIORETO"] = indicadoresReto.COSTO;
            dt.Rows[32]["PRECIOMSRETO"] = indicadoresReto.PRECIOKGMS;

            DataTable dtIxA = new DataTable();
            string query = @"SELECT (T1.LecProduc * T1.PrecLEc) / T2.IT                                   AS IxA
                                    FROM
                                    (
	                                    SELECT  '1'                                                       AS UNIR
	                                           ,SUM(med_lecproduc)                                        AS LecProduc
	                                           ,IIF( SUM(med_lecproduc)> 0, SUM(med_precioleche * med_lecproduc) / SUM(med_lecproduc), 0) AS PrecLEc
	                                    FROM media
	                                    WHERE med_fecha BETWEEN @fechaI AND @fechaF
	                                    AND ran_id IN (@ranchos) 
                                    ) T1
                                    LEFT JOIN
                                    (
	                                    SELECT  '1'                                                                                                                    AS UNIR
	                                           ,SUM(ia_vacas_ord + ia_vacas_secas + ia_vqreto + ia_vcreto + ia_jaulas + ia_destetadas + ia_destetadas2 + ia_vaquillas) AS IT
	                                    FROM inventario_afi
	                                    WHERE ia_fecha BETWEEN @fechaI AND @fechaF
	                                    AND ran_id IN (@ranchos)
                                    ) T2
                                    ON T1.UNIR = T2.UNIR";
            query = query.Replace("@fechaI", "'" + Finicio.AddDays(1).Date.ToString("yyyy-MM-dd HH:mm") + "'")
                         .Replace("@fechaF", "'" + FFinal.Date.ToString("yyyy-MM-dd HH:mm") + "'")
                         .Replace("@ranchos", ran_id.ToString());
            conn.QueryAlimento(query, out dtIxA);
            double producion = Convert.ToDouble(TablaGlobal.Rows[32][2]) * Convert.ToDouble(TablaGlobal.Rows[32]["PRECIOPROD"]);
            double jaulas = indicadoresJaulas.ANIMELES * indicadoresJaulas.COSTO;
            double destetadasUno = indicadoresDesteteUno.ANIMELES * indicadoresDesteteUno.COSTO;
            double destetadasDos = indicadoresDesteteDos.ANIMELES * indicadoresDesteteDos.COSTO;
            double vaquillas = indicadoresVaquillas.ANIMELES * indicadoresVaquillas.COSTO;
            double secas = indicadoresSecas.ANIMELES * indicadoresSecas.COSTO;
            double reto = indicadoresReto.ANIMELES * indicadoresReto.COSTO;
            double costo = Convert.ToDouble(TablaGlobal.Rows[32][2]) + indicadoresJaulas.ANIMELES + indicadoresDesteteUno.ANIMELES +
                           indicadoresDesteteDos.ANIMELES + indicadoresVaquillas.ANIMELES + indicadoresSecas.ANIMELES + indicadoresReto.ANIMELES;
            double todosjuntos = producion + jaulas + destetadasUno + destetadasDos + vaquillas + secas + reto;
            dt.Rows[32][44] = dtIxA.Rows[0][0] != DBNull.Value ? Convert.ToDouble(dtIxA.Rows[0][0]) : 0;
            dt.Rows[32][45] = todosjuntos / costo;
            dt.Rows[32][46] = (Convert.ToDouble(dt.Rows[32][45]) / Convert.ToDouble(dt.Rows[32][44])) * 100;
            dt.Rows[32][48] = Convert.ToDouble(dt.Rows[32][44]) - Convert.ToDouble(dt.Rows[32][45]);
            dt.Rows[32][49] = (Convert.ToDouble(dt.Rows[32][48]) / Convert.ToDouble(dt.Rows[32][44])) * 100;
        }

        public void Hora_Corte(out int horas, out int hcorte)
        {
            DataTable dt;
            string query = "select paramvalue from bedrijf_params where name = 'DSTimeShift' ";
            conn.QueryTracker(query, out dt);

            horas = Convert.ToInt32(dt.Rows[0][0]);
            hcorte = horas > 0 ? horas : 24 + horas;
        }

        #region ajustes con colorimetria
        private DataTable DtColorometriaHoja1()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("DIA").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("TOTALVTA").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("ORDENO").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("SECAS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("HATO").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCLACT").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCPROT").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("UREA").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCGRA").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("CCS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("CTD").DataType = System.Type.GetType("System.Double");
            //PROD
            dt.Columns.Add("LECHE").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("ANTIB").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("X").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("TOTALPROD").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("DEL").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("ANT").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("METAPROD").DataType = System.Type.GetType("System.Double");
            //VENTA
            dt.Columns.Add("ILCAVTA").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("ICVTA").DataType = System.Type.GetType("System.Double");
            //ALIMENTACION PRODUCCION
            dt.Columns.Add("EA").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("ILCAPROD").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("ICPROD").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIOL").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("MH").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCMS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("MS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("SA").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("MSS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("EAS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIOPROD").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIOMS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("METAMS").DataType = System.Type.GetType("System.Double");
            //CRIBAS
            dt.Columns.Add("N1").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("N2").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("N3").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("N4").DataType = System.Type.GetType("System.Double");
            //NO ID
            dt.Columns.Add("SES1").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("SES2").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("SES3").DataType = System.Type.GetType("System.Int32");
            //DEC 
            dt.Columns.Add("DEC").DataType = System.Type.GetType("System.Double");
            //Colores
            dt.Columns.Add("COLOR_ILCA").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_IC").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_PRECIOL").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_MH").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_PORCMS").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_MS").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_COSTOPROD").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_PRECIOKGMS").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_LECHE").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_MEDIA").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_TOTAL").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_SES1").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_SES2").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_SES3").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_CTD").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_ANT").DataType = System.Type.GetType("System.String");
            /*Columnas nuevas
            dt.Columns.Add("COLOR_MSS").DataType = System.Type.GetType("System.String");
            */

            return dt;
        }
        private DataTable DtColorometriaHoja2()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("DIA").DataType = System.Type.GetType("System.String");
            //Jaulas
            dt.Columns.Add("INV").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("PRECIOJ").DataType = System.Type.GetType("System.Double");
            // 2/7
            dt.Columns.Add("INV2").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("MH2").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIO2").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCMS2").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("MS2").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIOMS2").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("METABEC1").DataType = System.Type.GetType("System.Double");
            // 7/13
            dt.Columns.Add("INV7").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("MH7").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIO7").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCMS7").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("MS7").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIOMS7").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("METABEC2").DataType = System.Type.GetType("System.Double");
            // 13 a Mas
            dt.Columns.Add("INV13").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("MH13").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIO13").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCMS13").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("MS13").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIOMS13").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("METAVP").DataType = System.Type.GetType("System.Double");
            //Secas
            dt.Columns.Add("INVSECAS").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("MHSECAS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCMSSECAS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("MSSECAS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("SASECAS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("MSSSECAS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCSSECAS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIOSECAS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIOMSSECAS").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("METASECAS").DataType = System.Type.GetType("System.Double");
            //Reto
            dt.Columns.Add("INVRETO").DataType = System.Type.GetType("System.Int32");
            dt.Columns.Add("MHRETO").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCMSRETO").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("MSRETO").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("SARETO").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("MSSRETO").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCSRETO").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIORETO").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PRECIOMSRETO").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("METARETO").DataType = System.Type.GetType("System.Double");
            //Utilidad Por Animal
            dt.Columns.Add("IXA").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("CXA").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCENTAJE1").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("IT").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("UXA").DataType = System.Type.GetType("System.Double");
            dt.Columns.Add("PORCENTAJE2").DataType = System.Type.GetType("System.Double");
            //Colores            
            dt.Columns.Add("COLOR_MH_CRECIMIENTO").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_PORCMS_CRECIMIENTO").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_MS_CRECIMIENTO").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_PRECIOKGMS_CRECIMIENTO").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_COSTO_CRECIMIENTO").DataType = System.Type.GetType("System.String");

            dt.Columns.Add("COLOR_MH_DESARROLLO").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_PORCMS_DESARROLLO").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_MS_DESARROLLO").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_PRECIOKGMS_DESARROLLO").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_COSTO_DESARROLLO").DataType = System.Type.GetType("System.String");

            dt.Columns.Add("COLOR_MH_VAQUILLAS").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_PORCMS_VAQUILLAS").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_MS_VAQUILLAS").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_PRECIOKGMS_VAQUILLAS").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_COSTO_VAQUILLAS").DataType = System.Type.GetType("System.String");

            dt.Columns.Add("COLOR_MH_SECAS").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_PORCMS_SECAS").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_MS_SECAS").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_PRECIOKGMS_SECAS").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_COSTO_SECAS").DataType = System.Type.GetType("System.String");

            dt.Columns.Add("COLOR_MH_RETO").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_PORCMS_RETO").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_MS_RETO").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_PRECIOKGMS_RETO").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_COSTO_RETO").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_IXA").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_CXA").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_PORCENTAJEC").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_UXA").DataType = System.Type.GetType("System.String");
            dt.Columns.Add("COLOR_PORCENTAJEU").DataType = System.Type.GetType("System.String");

            return dt;
        }

        private DataTable DTHoja1ConColorimetria(DataTable dtHoja1)
        {
            DataTable dt = DtColorometriaHoja1();


            decimal media = 0, antib = 0, ant = 0; 
            foreach (DataRow row in dtHoja1.Rows)
            {
                DataRow newRow = dt.NewRow();
                newRow["DIA"] = row["DIA"];
                newRow["TOTALVTA"] = row["TOTALVTA"];
                newRow["ORDENO"] = row["ORDENO"];
                newRow["SECAS"] = row["SECAS"];
                newRow["HATO"] = row["HATO"];
                newRow["PORCPROT"] = row["PORCPROT"];
                newRow["PORCLACT"] = row["PORCLACT"];
                newRow["UREA"] = row["UREA"];
                newRow["PORCGRA"] = row["PORCGRA"];
                newRow["CCS"] = row["CCS"];
                newRow["CTD"] = row["CTD"];
                newRow["LECHE"] = row["LECHE"];
                newRow["ANTIB"] = row["ANTIB"];
                newRow["X"] = row["X"];
                newRow["TOTALPROD"] = row["TOTALPROD"];
                newRow["DEL"] = row["DEL"];
                newRow["ANT"] = row["ANT"];
                newRow["METAPROD"] = row["METAPROD"];
                newRow["ILCAVTA"] = row["ILCAVTA"];
                newRow["ICVTA"] = row["ICVTA"];
                newRow["EA"] = row["EA"];
                newRow["ILCAPROD"] = row["ILCAPROD"];
                newRow["ICPROD"] = row["ICPROD"];
                newRow["PRECIOL"] = row["PRECIOL"];
                newRow["MH"] = row["MH"];
                newRow["PORCMS"] = row["PORCMS"];
                newRow["MS"] = row["MS"];
                newRow["SA"] = row["SA"];
                newRow["MSS"] = row["MSS"];
                newRow["EAS"] = row["EAS"];
                newRow["PORCS"] = row["PORCS"];
                newRow["PRECIOPROD"] = row["PRECIOPROD"];
                newRow["PRECIOMS"] = row["PRECIOMS"];
                newRow["METAMS"] = row["METAMS"];
                newRow["N1"] = row["N1"];
                newRow["N2"] = row["N2"];
                newRow["N3"] = row["N3"];
                newRow["N4"] = row["N4"];
                newRow["SES1"] = row["SES1"];
                newRow["SES2"] = row["SES2"];
                newRow["SES3"] = row["SES3"];
                newRow["DEC"] = row["DEC"];
                //colores
                if (row["DIA"].ToString() != "TOTAL" && row["DIA"].ToString() != "PROM" && row["DIA"].ToString() != "AÑO ANT" &&
                    row["DIA"].ToString() != "DIF %" && row["DIA"].ToString() != "DIF #" && row["DIA"].ToString() != "NA")
                {
                    if (row["ILCAPROD"] != DBNull.Value)
                        newRow["COLOR_ILCA"] = Color1Hoja1(Convert.ToDecimal(row["ILCAPROD"].ToString()), Convert.ToDecimal(prom_ordeño.ILCA_PRODUCCION));
                    else
                        newRow["COLOR_ILCA"] = "";

                    if (row["ICPROD"] != DBNull.Value)
                        newRow["COLOR_IC"] = Color1Hoja1(Convert.ToDecimal(row["ICPROD"].ToString()), Convert.ToDecimal(prom_ordeño.IC_PRODUCCION));
                    else
                        newRow["COLOR_IC"] = "";

                    if (row["PRECIOL"] != DBNull.Value)
                        newRow["COLOR_PRECIOL"] = Color2Hoja1(Convert.ToDecimal(row["PRECIOL"].ToString()), Convert.ToDecimal(prom_ordeño.PRECIOL));
                    else
                        newRow["COLOR_PRECIOL"] = "";

                    if (row["LECHE"] != DBNull.Value)
                        newRow["COLOR_LECHE"] = Color1Hoja1(Convert.ToDecimal(row["LECHE"].ToString()), datosProd.LECHE);
                    else
                        newRow["COLOR_LECHE"] = "";

                    if (row["X"] != DBNull.Value)
                    {
                        newRow["COLOR_MEDIA"] = Color1Hoja1(Convert.ToDecimal(row["X"].ToString()), datosProd.MEDIA);
                        media = Convert.ToDecimal(row["X"]);
                    }
                    else
                        newRow["COLOR_MEDIA"] = "";

                    if (row["TOTALPROD"] != DBNull.Value)
                        newRow["COLOR_TOTAL"] = Color1Hoja1(Convert.ToDecimal(row["TOTALPROD"].ToString()), datosProd.TOTAL);
                    else
                        newRow["COLOR_TOTAL"] = "";

                    if (row["ANT"] != DBNull.Value)
                    {
                        if (row["ANTIB"] != DBNull.Value)
                            antib = Convert.ToDecimal(row["ANTIB"].ToString());
                        else
                            antib = 0;

                        ant = Convert.ToDecimal(row["ANT"]);
                        newRow["COLOR_ANT"] = ColorAntHoja1(media, antib, ant);
                    }
                    else
                        newRow["COLOR_ANT"] = "";

                    if (row["SES1"] != DBNull.Value)
                        newRow["COLOR_SES1"] = ColorSes(Convert.ToInt32(row["ORDENO"]), Convert.ToInt32(row["SES1"]));
                    else
                        newRow["COLOR_SES1"] = "";

                    if (row["SES2"] != DBNull.Value)
                        newRow["COLOR_SES2"] = ColorSes(Convert.ToInt32(row["ORDENO"]), Convert.ToInt32(row["SES2"]));
                    else
                        newRow["COLOR_SES2"] = "";

                    if (row["SES3"] != DBNull.Value)
                        newRow["COLOR_SES3"] = ColorSes(Convert.ToInt32(row["ORDENO"]), Convert.ToInt32(row["SES3"]));
                    else
                        newRow["COLOR_SES3"] = "";

                    if (row["DIA"] != DBNull.Value && row["DIA"].ToString() != "")
                    {
                        DateTime dia = new DateTime(fechaFinDTP.Year, fechaFinDTP.Month, Convert.ToInt32(row["DIA"].ToString()));
                        List<DatosTeorico> busquedaProd = (from x in listaProduccion where x.FECHA == dia select x).ToList();
                        List<LecheBacteriologia> busquedaBacteorologia = (from x in listaBacteriologia where x.FECHA == dia select x).ToList();

                        if (busquedaProd.Count > 0)
                        {
                            decimal invOrdeño = 0;
                            decimal.TryParse(row["ORDENO"].ToString(), out invOrdeño);
                            if (row["MH"].ToString() != "")
                                newRow["COLOR_MH"] = ColorDato(invOrdeño, Convert.ToDecimal(row["MH"].ToString()), busquedaProd[0].MH);
                            else
                                newRow["COLOR_MH"] = ColorDato(invOrdeño, 0, busquedaProd[0].MH); ;

                            if (row["PORCMS"].ToString() != "")
                                newRow["COLOR_PORCMS"] = ColorDato(invOrdeño, Convert.ToDecimal(row["PORCMS"].ToString()), busquedaProd[0].PORCENTAJE_MS);
                            else
                                newRow["COLOR_PORCMS"] = ColorDato(invOrdeño, 0, busquedaProd[0].PORCENTAJE_MS);

                            if (row["MS"].ToString() != "")
                                newRow["COLOR_MS"] = ColorDato(invOrdeño, Convert.ToDecimal(row["MS"].ToString()), busquedaProd[0].MS);
                            else
                                newRow["COLOR_MS"] = ColorDato(invOrdeño, 0, busquedaProd[0].MS);

                            if (row["PRECIOPROD"].ToString() != "")
                                newRow["COLOR_COSTOPROD"] = ColorDato(invOrdeño, Convert.ToDecimal(row["PRECIOPROD"].ToString()), busquedaProd[0].COSTO);
                            else
                                newRow["COLOR_COSTOPROD"] = ColorDato(invOrdeño, 0, busquedaProd[0].COSTO);

                            if (row["PRECIOMS"].ToString() != "")
                                newRow["COLOR_PRECIOKGMS"] = ColorDato(invOrdeño, Convert.ToDecimal(row["PRECIOMS"].ToString()), busquedaProd[0].KGMS);
                            else
                                newRow["COLOR_PRECIOKGMS"] = ColorDato(invOrdeño, 0, busquedaProd[0].KGMS);

                        }
                        else
                        {
                            newRow["COLOR_ILCA"] = "";
                            newRow["COLOR_IC"] = "";
                            newRow["COLOR_PRECIOL"] = "";
                            newRow["COLOR_MH"] = "";
                            newRow["COLOR_PORCMS"] = "";
                            newRow["COLOR_MS"] = "";
                            newRow["COLOR_COSTOPROD"] = "";
                            newRow["COLOR_PRECIOKGMS"] = "";
                        }

                        if (!(newRow["CTD"].ToString() == ""))
                        {
                            if (busquedaBacteorologia.Count > 0)
                            {
                                if (busquedaBacteorologia[0].BACTEROLOGIA < 0.2M)
                                    newRow["COLOR_CTD"] = "#FFC9C9";
                                else
                                    newRow["COLOR_CTD"] = "#F2F2F2";
                            }
                            else
                                newRow["COLOR_CTD"] = "";
                        }
                        else
                            newRow["COLOR_CTD"] = "";

                    }



                }
                else
                {
                    newRow["COLOR_ILCA"] = "";
                    newRow["COLOR_IC"] = "";
                    newRow["COLOR_PRECIOL"] = "";
                    newRow["COLOR_MH"] = "";
                    newRow["COLOR_PORCMS"] = "";
                    newRow["COLOR_MS"] = "";
                    newRow["COLOR_COSTOPROD"] = "";
                    newRow["COLOR_PRECIOKGMS"] = "";
                    newRow["COLOR_LECHE"] = "";
                    newRow["COLOR_MEDIA"] = "";
                    newRow["COLOR_TOTAL"] = "";
                    newRow["COLOR_CTD"] = "";
                    newRow["COLOR_SES1"] = "";
                    newRow["COLOR_SES2"] = "";
                    newRow["COLOR_SES3"] = "";
                    newRow["COLOR_ANT"] = "";
                }
                dt.Rows.Add(newRow);
            }

            return dt;
        }

        private string ColorAntHoja1(decimal media, decimal antib, decimal ant)
        {
            string color = "";
            try 
            {
                decimal result = ant > 0 ? antib / ant : 0;

                switch(result)
                {
                    //Color Verde
                    case decimal n when (n >= (media * 0.5M) && n <= (media * 0.7M)):
                        color = "#DEEDD3";
                        break;
                    //Color Amarillo
                    case decimal n when ((n > (media * 0.7M) && n <= (media * 0.8M)) || (n >= (media * 0.4M) && n < (media * 0.5M)) ):
                        color = "#FFF5D9";
                        break;
                    //Color Rojo
                    default:
                        color = "#FFC9C9";
                        break;
                }
            
            }
            catch { }

            return color;
        }

        private DataTable DTHoja2ConColorimetria(DataTable dtHoja2)
        {
            DataTable dt = DtColorometriaHoja2();

            foreach (DataRow row in dtHoja2.Rows)
            {
                DataRow newRow = dt.NewRow();
                newRow["DIA"] = row["DIA"];
                newRow["INV"] = row["INV"];
                newRow["PRECIOJ"] = row["PRECIOJ"];
                newRow["INV2"] = row["INV2"];
                newRow["MH2"] = row["MH2"];
                newRow["PRECIO2"] = row["PRECIO2"];
                newRow["PORCMS2"] = row["PORCMS2"];
                newRow["MS2"] = row["MS2"];
                newRow["PRECIOMS2"] = row["PRECIOMS2"];
                newRow["METABEC1"] = row["METABEC1"];
                newRow["INV7"] = row["INV7"];
                newRow["MH7"] = row["MH7"];
                newRow["PRECIO7"] = row["PRECIO7"];
                newRow["PORCMS7"] = row["PORCMS7"];
                newRow["MS7"] = row["MS7"];
                newRow["PRECIOMS7"] = row["PRECIOMS7"];
                newRow["METABEC2"] = row["METABEC2"];
                newRow["INV13"] = row["INV13"];
                newRow["MH13"] = row["MH13"];
                newRow["PRECIO13"] = row["PRECIO13"];
                newRow["PORCMS13"] = row["PORCMS13"]; //Covenant lucido cambio los campos               
                newRow["MS13"] = row["MS13"];
                newRow["PRECIOMS13"] = row["PRECIOMS13"]; //Covenant lucido cambio los campos
                newRow["METAVP"] = row["METAVP"];
                newRow["INVSECAS"] = row["INVSECAS"];
                newRow["MHSECAS"] = row["MHSECAS"];
                newRow["PORCMSSECAS"] = row["PORCMSSECAS"];
                newRow["MSSECAS"] = row["MSSECAS"];
                newRow["SASECAS"] = row["SASECAS"];
                newRow["MSSSECAS"] = row["MSSSECAS"];
                newRow["PORCSSECAS"] = row["PORCSSECAS"];
                newRow["PRECIOSECAS"] = row["PRECIOSECAS"];
                newRow["PRECIOMSSECAS"] = row["PRECIOMSSECAS"];
                newRow["METASECAS"] = row["METASECAS"];
                newRow["INVRETO"] = row["INVRETO"];
                newRow["MHRETO"] = row["MHRETO"];
                newRow["PORCMSRETO"] = row["PORCMSRETO"];
                newRow["MSRETO"] = row["MSRETO"];
                newRow["SARETO"] = row["SARETO"];
                newRow["MSSRETO"] = row["MSSRETO"];
                newRow["PORCSRETO"] = row["PORCSRETO"];
                newRow["PRECIORETO"] = row["PRECIORETO"];
                newRow["PRECIOMSRETO"] = row["PRECIOMSRETO"];
                newRow["METARETO"] = row["METARETO"];
                newRow["IXA"] = row["IXA"];
                newRow["CXA"] = row["CXA"];
                newRow["PORCENTAJE1"] = row["PORCENTAJE1"];
                newRow["IT"] = row["IT"];
                newRow["UXA"] = row["UXA"];
                newRow["PORCENTAJE2"] = row["PORCENTAJE2"];

                if (row["DIA"].ToString() != "TOTAL" && row["DIA"].ToString() != "PROM" && row["DIA"].ToString() != "AÑO ANT" &&
                    row["DIA"].ToString() != "DIF %" && row["DIA"].ToString() != "DIF #" && row["DIA"].ToString() != "NA")
                {
                    if (row["DIA"] != DBNull.Value && row["DIA"].ToString() != "")
                    {
                        DateTime dia = new DateTime(fechaFinDTP.Year, fechaFinDTP.Month, Convert.ToInt32(row["DIA"].ToString()));
                        List<DatosTeorico> busquedaCrecimiento = (from x in listaDestet1 where x.FECHA == dia select x).ToList();
                        List<DatosTeorico> busquedaDesarrollo = (from x in listaDestete2 where x.FECHA == dia select x).ToList();
                        List<DatosTeorico> busquedaVaquillas = (from x in listaVaquillas where x.FECHA == dia select x).ToList();
                        List<DatosTeorico> busquedaSecas = (from x in listaSecas where x.FECHA == dia select x).ToList();
                        List<DatosTeorico> busquedaReto = (from x in listaReto where x.FECHA == dia select x).ToList();

                        if (busquedaCrecimiento.Count > 0)
                        {
                            decimal invCrecimiento = 0; decimal valorMH2 = 0; decimal valorMS2 = 0;
                            decimal.TryParse(row["INV2"].ToString(), out invCrecimiento);
                            decimal.TryParse(row["MH2"].ToString(), out valorMH2);
                            decimal.TryParse(row["MS2"].ToString(), out valorMS2);
                            newRow["COLOR_MH_CRECIMIENTO"] = row["MH2"] != DBNull.Value ? ColorDato(invCrecimiento, valorMH2, busquedaCrecimiento[0].MH) : ColorDato(invCrecimiento, 0, busquedaCrecimiento[0].MH);
                            newRow["COLOR_PORCMS_CRECIMIENTO"] = row["PORCMS2"] != DBNull.Value ? ColorDato(invCrecimiento, Convert.ToDecimal(row["PORCMS2"].ToString()), busquedaCrecimiento[0].PORCENTAJE_MS) : ColorDato(invCrecimiento, 0, busquedaCrecimiento[0].PORCENTAJE_MS);
                            newRow["COLOR_MS_CRECIMIENTO"] = row["MS2"] != DBNull.Value ? ColorDato(invCrecimiento, valorMS2, busquedaCrecimiento[0].MS) : ColorDato(invCrecimiento, 0, busquedaCrecimiento[0].MS);
                            newRow["COLOR_PRECIOKGMS_CRECIMIENTO"] = row["PRECIOMS2"] != DBNull.Value ? ColorDato(invCrecimiento, Convert.ToDecimal(row["PRECIOMS2"].ToString()), busquedaCrecimiento[0].KGMS) : ColorDato(invCrecimiento, 0, busquedaCrecimiento[0].KGMS);
                            newRow["COLOR_COSTO_CRECIMIENTO"] = row["PRECIO2"] != DBNull.Value ? ColorDato(invCrecimiento, Convert.ToDecimal(row["PRECIO2"].ToString()), busquedaCrecimiento[0].COSTO) : ColorDato(invCrecimiento, 0, busquedaCrecimiento[0].COSTO);
                        }
                        else
                        {
                            newRow["COLOR_MH_CRECIMIENTO"] = "";
                            newRow["COLOR_PORCMS_CRECIMIENTO"] = "";
                            newRow["COLOR_MS_CRECIMIENTO"] = "";
                            newRow["COLOR_PRECIOKGMS_CRECIMIENTO"] = "";
                            newRow["COLOR_COSTO_CRECIMIENTO"] = "";
                        }

                        if (busquedaDesarrollo.Count > 0)
                        {
                            decimal invDesarrollo = 0;
                            decimal.TryParse(row["INV7"].ToString(), out invDesarrollo);
                            newRow["COLOR_MH_DESARROLLO"] = row["MH7"] != DBNull.Value ? ColorDato(invDesarrollo, Convert.ToDecimal(row["MH7"].ToString()), busquedaDesarrollo[0].MH) : ColorDato(invDesarrollo, 0, busquedaDesarrollo[0].MH);
                            newRow["COLOR_PORCMS_DESARROLLO"] = row["PORCMS7"] != DBNull.Value ? ColorDato(invDesarrollo, Convert.ToDecimal(row["PORCMS7"].ToString()), busquedaDesarrollo[0].PORCENTAJE_MS) : ColorDato(invDesarrollo, 0, busquedaDesarrollo[0].PORCENTAJE_MS);
                            newRow["COLOR_MS_DESARROLLO"] = row["MS7"] != DBNull.Value ? ColorDato(invDesarrollo, Convert.ToDecimal(row["MS7"].ToString()), busquedaDesarrollo[0].MS) : ColorDato(invDesarrollo, 0, busquedaDesarrollo[0].MS);
                            newRow["COLOR_PRECIOKGMS_DESARROLLO"] = row["PRECIOMS7"] != DBNull.Value ? ColorDato(invDesarrollo, Convert.ToDecimal(row["PRECIOMS7"].ToString()), busquedaDesarrollo[0].KGMS) : ColorDato(invDesarrollo, 0, busquedaDesarrollo[0].KGMS);
                            newRow["COLOR_COSTO_DESARROLLO"] = row["PRECIO7"] != DBNull.Value ? ColorDato(invDesarrollo, Convert.ToDecimal(row["PRECIO7"].ToString()), busquedaDesarrollo[0].COSTO) : ColorDato(invDesarrollo, 0, busquedaDesarrollo[0].COSTO);
                        }
                        else
                        {
                            newRow["COLOR_MH_DESARROLLO"] = "";
                            newRow["COLOR_PORCMS_DESARROLLO"] = "";
                            newRow["COLOR_MS_DESARROLLO"] = "";
                            newRow["COLOR_PRECIOKGMS_DESARROLLO"] = "";
                            newRow["COLOR_COSTO_DESARROLLO"] = "";
                        }

                        if (busquedaVaquillas.Count > 0)
                        {
                            decimal invVaquillas = 0;
                            decimal.TryParse(row["INV13"].ToString(), out invVaquillas);
                            newRow["COLOR_MH_VAQUILLAS"] = row["MH13"] != DBNull.Value ? ColorDato(invVaquillas, Convert.ToDecimal(row["MH13"].ToString()), busquedaVaquillas[0].MH) : ColorDato(invVaquillas, 0, busquedaVaquillas[0].MH);
                            newRow["COLOR_PORCMS_VAQUILLAS"] = row["PORCMS13"] != DBNull.Value ? ColorDato(invVaquillas, Convert.ToDecimal(row["PORCMS13"].ToString()), busquedaVaquillas[0].PORCENTAJE_MS) : ColorDato(invVaquillas, 0, busquedaVaquillas[0].PORCENTAJE_MS);
                            newRow["COLOR_MS_VAQUILLAS"] = row["MS13"] != DBNull.Value ? ColorDato(invVaquillas, Convert.ToDecimal(row["MS13"].ToString()), busquedaVaquillas[0].MS) : ColorDato(invVaquillas, 0, busquedaVaquillas[0].MS);
                            newRow["COLOR_PRECIOKGMS_VAQUILLAS"] = row["PRECIOMS13"] != DBNull.Value ? ColorDato(invVaquillas, Convert.ToDecimal(row["PRECIOMS13"].ToString()), busquedaVaquillas[0].KGMS) : ColorDato(invVaquillas, 0, busquedaVaquillas[0].KGMS);
                            newRow["COLOR_COSTO_VAQUILLAS"] = row["PRECIO13"] != DBNull.Value ? ColorDato(invVaquillas, Convert.ToDecimal(row["PRECIO13"].ToString()), busquedaVaquillas[0].COSTO) : ColorDato(invVaquillas, 0, busquedaVaquillas[0].COSTO);
                        }
                        else
                        {
                            newRow["COLOR_MH_VAQUILLAS"] = "";
                            newRow["COLOR_PORCMS_VAQUILLAS"] = "";
                            newRow["COLOR_MS_VAQUILLAS"] = "";
                            newRow["COLOR_PRECIOKGMS_VAQUILLAS"] = "";
                            newRow["COLOR_COSTO_VAQUILLAS"] = "";
                        }

                        if (busquedaSecas.Count > 0)
                        {
                            decimal invSecas = 0;
                            decimal.TryParse(row["INVSECAS"].ToString(), out invSecas);
                            newRow["COLOR_MH_SECAS"] = row["MHSECAS"] != DBNull.Value ? ColorDato(invSecas, Convert.ToDecimal(row["MHSECAS"].ToString()), busquedaSecas[0].MH) : ColorDato(invSecas, 0, busquedaSecas[0].MH);
                            newRow["COLOR_PORCMS_SECAS"] = row["PORCMSSECAS"] != DBNull.Value ? ColorDato(invSecas, Convert.ToDecimal(row["PORCMSSECAS"].ToString()), busquedaSecas[0].PORCENTAJE_MS) : ColorDato(invSecas, 0, busquedaSecas[0].PORCENTAJE_MS);
                            newRow["COLOR_MS_SECAS"] = row["MSSECAS"] != DBNull.Value ? ColorDato(invSecas, Convert.ToDecimal(row["MSSECAS"].ToString()), busquedaSecas[0].MS) : ColorDato(invSecas, 0, busquedaSecas[0].MS);
                            newRow["COLOR_PRECIOKGMS_SECAS"] = row["PRECIOMSSECAS"] != DBNull.Value ? ColorDato(invSecas, Convert.ToDecimal(row["PRECIOMSSECAS"].ToString()), busquedaSecas[0].KGMS) : ColorDato(invSecas, 0, busquedaSecas[0].KGMS);
                            newRow["COLOR_COSTO_SECAS"] = row["PRECIOSECAS"] != DBNull.Value ? ColorDato(invSecas, Convert.ToDecimal(row["PRECIOSECAS"].ToString()), busquedaSecas[0].COSTO) : ColorDato(invSecas, 0, busquedaSecas[0].COSTO);
                        }
                        else
                        {
                            newRow["COLOR_MH_SECAS"] = "";
                            newRow["COLOR_PORCMS_SECAS"] = "";
                            newRow["COLOR_MS_SECAS"] = "";
                            newRow["COLOR_PRECIOKGMS_SECAS"] = "";
                            newRow["COLOR_COSTO_SECAS"] = "";
                        }

                        if (busquedaReto.Count > 0)
                        {
                            decimal invReto = 0;
                            decimal.TryParse(row["INVRETO"].ToString(), out invReto);
                            newRow["COLOR_MH_RETO"] = row["MHRETO"] != DBNull.Value ? ColorDato(invReto, Convert.ToDecimal(row["MHRETO"].ToString()), busquedaReto[0].MH) : ColorDato(invReto, 0, busquedaReto[0].MH);
                            newRow["COLOR_PORCMS_RETO"] = row["PORCMSRETO"] != DBNull.Value ? ColorDato(invReto, Convert.ToDecimal(row["PORCMSRETO"].ToString()), busquedaReto[0].PORCENTAJE_MS) : ColorDato(invReto, 0, busquedaReto[0].PORCENTAJE_MS);
                            newRow["COLOR_MS_RETO"] = row["MSRETO"] != DBNull.Value ? ColorDato(invReto, Convert.ToDecimal(row["MSRETO"].ToString()), busquedaReto[0].MS) : ColorDato(invReto, 0, busquedaReto[0].MS);
                            newRow["COLOR_PRECIOKGMS_RETO"] = row["PRECIOMSRETO"] != DBNull.Value ? ColorDato(invReto, Convert.ToDecimal(row["PRECIOMSRETO"].ToString()), busquedaReto[0].KGMS) : ColorDato(invReto, 0, busquedaReto[0].KGMS);
                            newRow["COLOR_COSTO_RETO"] = row["PRECIORETO"] != DBNull.Value ? ColorDato(invReto, Convert.ToDecimal(row["PRECIORETO"].ToString()), busquedaReto[0].COSTO) : ColorDato(invReto, 0, busquedaReto[0].COSTO);
                        }
                        else
                        {
                            newRow["COLOR_MH_RETO"] = "";
                            newRow["COLOR_PORCMS_RETO"] = "";
                            newRow["COLOR_MS_RETO"] = "";
                            newRow["COLOR_PRECIOKGMS_RETO"] = "";
                            newRow["COLOR_COSTO_RETO"] = "";
                        }

                        if (row["IXA"] != DBNull.Value)
                            newRow["COLOR_IXA"] = ColorUtilidad(Convert.ToDecimal(row["IXA"].ToString()), utilidad.IXA, false);
                        else
                            newRow["COLOR_IXA"] = "";

                        if (row["CXA"] != DBNull.Value)
                            newRow["COLOR_CXA"] = ColorUtilidad(Convert.ToDecimal(row["CXA"].ToString()), utilidad.CXA, true);
                        else
                            newRow["COLOR_CXA"] = "";

                        if (row["PORCENTAJE1"] != DBNull.Value)
                            newRow["COLOR_PORCENTAJEC"] = ColorUtilidad(Convert.ToDecimal(row["PORCENTAJE1"].ToString()), utilidad.PORCENTAJE_C, false);
                        else
                            newRow["COLOR_PORCENTAJEC"] = "";

                        if (row["UXA"] != DBNull.Value)
                            newRow["COLOR_UXA"] = ColorUtilidad(Convert.ToDecimal(row["UXA"].ToString()), utilidad.UXA, false);
                        else
                            newRow["COLOR_UXA"] = "";

                        if (row["PORCENTAJE2"] != DBNull.Value)
                            newRow["COLOR_PORCENTAJEU"] = ColorUtilidad(Convert.ToDecimal(row["PORCENTAJE2"].ToString()), utilidad.PORCENTAJE_U, false);
                        else
                            newRow["COLOR_PORCENTAJEU"] = "";
                    }
                    else
                    {
                        newRow["COLOR_MH_CRECIMIENTO"] = "";
                        newRow["COLOR_PORCMS_CRECIMIENTO"] = "";
                        newRow["COLOR_MS_CRECIMIENTO"] = "";
                        newRow["COLOR_PRECIOKGMS_CRECIMIENTO"] = "";
                        newRow["COLOR_COSTO_CRECIMIENTO"] = "";

                        newRow["COLOR_MH_DESARROLLO"] = "";
                        newRow["COLOR_PORCMS_DESARROLLO"] = "";
                        newRow["COLOR_MS_DESARROLLO"] = "";
                        newRow["COLOR_PRECIOKGMS_DESARROLLO"] = "";
                        newRow["COLOR_COSTO_DESARROLLO"] = "";

                        newRow["COLOR_MH_VAQUILLAS"] = "";
                        newRow["COLOR_PORCMS_VAQUILLAS"] = "";
                        newRow["COLOR_MS_VAQUILLAS"] = "";
                        newRow["COLOR_PRECIOKGMS_VAQUILLAS"] = "";
                        newRow["COLOR_COSTO_VAQUILLAS"] = "";

                        newRow["COLOR_MH_SECAS"] = "";
                        newRow["COLOR_PORCMS_SECAS"] = "";
                        newRow["COLOR_MS_SECAS"] = "";
                        newRow["COLOR_PRECIOKGMS_SECAS"] = "";
                        newRow["COLOR_COSTO_SECAS"] = "";

                        newRow["COLOR_MH_RETO"] = "";
                        newRow["COLOR_PORCMS_RETO"] = "";
                        newRow["COLOR_MS_RETO"] = "";
                        newRow["COLOR_PRECIOKGMS_RETO"] = "";
                        newRow["COLOR_COSTO_RETO"] = "";

                        newRow["COLOR_IXA"] = "";
                        newRow["COLOR_CXA"] = "";
                        newRow["COLOR_PORCENTAJEC"] = "";
                        newRow["COLOR_UXA"] = "";
                        newRow["COLOR_PORCENTAJEU"] = "";
                    }
                }
                else
                {
                    newRow["COLOR_MH_CRECIMIENTO"] = "";
                    newRow["COLOR_PORCMS_CRECIMIENTO"] = "";
                    newRow["COLOR_MS_CRECIMIENTO"] = "";
                    newRow["COLOR_PRECIOKGMS_CRECIMIENTO"] = "";

                    newRow["COLOR_MH_DESARROLLO"] = "";
                    newRow["COLOR_PORCMS_DESARROLLO"] = "";
                    newRow["COLOR_MS_DESARROLLO"] = "";
                    newRow["COLOR_PRECIOKGMS_DESARROLLO"] = "";

                    newRow["COLOR_MH_VAQUILLAS"] = "";
                    newRow["COLOR_PORCMS_VAQUILLAS"] = "";
                    newRow["COLOR_MS_VAQUILLAS"] = "";
                    newRow["COLOR_PRECIOKGMS_VAQUILLAS"] = "";

                    newRow["COLOR_MH_SECAS"] = "";
                    newRow["COLOR_PORCMS_SECAS"] = "";
                    newRow["COLOR_MS_SECAS"] = "";
                    newRow["COLOR_PRECIOKGMS_SECAS"] = "";

                    newRow["COLOR_MH_RETO"] = "";
                    newRow["COLOR_PORCMS_RETO"] = "";
                    newRow["COLOR_MS_RETO"] = "";
                    newRow["COLOR_PRECIOKGMS_RETO"] = "";

                    newRow["COLOR_IXA"] = "";
                    newRow["COLOR_CXA"] = "";
                    newRow["COLOR_PORCENTAJEC"] = "";
                    newRow["COLOR_UXA"] = "";
                    newRow["COLOR_PORCENTAJEU"] = "";
                }

                dt.Rows.Add(newRow);
            }

            return dt;
        }

        private string Color1Hoja1(decimal valor, decimal promedio)
        {
            string color = "";

            if (valor != 0)
            {
                switch (valor)
                {
                    //Blanco    #F2F2F2
                    case decimal n when ((n >= (promedio * 0.95M) && n <= (promedio * 1.05M))):
                        color = "#F2F2F2";
                        break;
                    //Verde #DEEDD3
                    case decimal n when (n > (promedio * 1.05M)):
                        color = "#DEEDD3";
                        break;
                    //Amarillo #FFF5D9
                    case decimal n when (n >= (promedio * 0.90M) && n < (promedio * 0.95M)):
                        color = "#FFF5D9";
                        break;
                    //Rojo  #FFC9C9
                    case decimal n when (n < (promedio * 0.90M)):
                        color = "#FFC9C9";
                        break;
                }
            }
            return color;
        }

        private string Color2Hoja1(decimal valor, decimal promedio)
        {
            string color = "";

            if (valor != 0)
            {
                switch (valor)
                {
                    //Verde #DEEDD3
                    case decimal n when (n < (promedio * 0.95M)):
                        color = "#DEEDD3";
                        break;
                    //Blanco    #F2F2F2
                    case decimal n when ((n >= (promedio * 0.95M) && n < (promedio * 1.05M))):
                        color = "#F2F2F2";
                        break;
                    //Amarillo #FFF5D9
                    case decimal n when (n >= (promedio * 1.05M) && n <= (promedio * 1.10M)):
                        color = "#FFF5D9";
                        break;
                    //Rojo  #FFC9C9
                    case decimal n when (n > (promedio * 1.10M)):
                        color = "#FFC9C9";
                        break;
                }
            }
            return color;
        }

        private string ColorDato(decimal inventario, decimal valor, decimal promedio)
        {
            string color = "";

            if (valor != 0)
            {
                switch (valor)
                {
                    //Color Verde
                    case decimal n when ((n >= (promedio * 0.95M) && n <= (promedio * 1.05M))):
                        color = "#DEEDD3";
                        break;
                    //Color Blanco
                    case decimal n when (n >= (promedio * 0.9M) && n < (promedio * 0.95M)):
                        color = "#F2F2F2";
                        break;
                    case decimal n when (n > (promedio * 1.05M) && n <= (promedio * 1.1M)):
                        color = "#F2F2F2";
                        break;
                    //Color Amarillo
                    case decimal n when (n >= (promedio * 0.85M) && n < (promedio * 0.9M)):
                        color = "#FFF5D9";
                        break;
                    case decimal n when (n > (promedio * 1.1M) && n <= (promedio * 1.15M)):
                        color = "#FFF5D9";
                        break;
                    //Color Rojo
                    case decimal n when (n < (promedio * 0.85M) || n > (promedio * 1.15M)):
                        color = "#FFC9C9";
                        break;
                }
            }
            else if (inventario != 0)
            {
                color = "#FFC9C9";
            }

            return color;
        }

        private string ColorSes(int inventario, int vacas_ses)
        {
            string color = "";

            if (pesadores)
            {
                if (vacas_ses != 0)
                {
                    decimal porcentaje = (vacas_ses * 1.0M) / inventario * 100;

                    switch (porcentaje)
                    {
                        case decimal n when (n >= 0 && n <= 3M):
                            color = "#DEEDD3";
                            break;
                        case decimal n when (n > 3M && n <= 5M):
                            color = "#F2F2F2";
                            break;
                        case decimal n when (n > 5M && n <= 8M):
                            color = "#FFF5D9";
                            break;
                        case decimal n when n > 8M:
                            color = "#FFC9C9";
                            break;
                    }

                }
            }
            return color;
        }

        private string ColorUtilidad(decimal valor, decimal promedio, bool esCosto)
        {
            string color = "";

            if (valor != 0)
            {
                if (esCosto)
                {
                    switch (valor)
                    {
                        //Verde #DEEDD3
                        case decimal n when (n < (promedio * 0.95M)):
                            color = "#DEEDD3";
                            break;
                        //Blanco    #F2F2F2
                        case decimal n when ((n >= (promedio * 0.95M) && n <= (promedio * 1.05M))):
                            color = "#F2F2F2";
                            break;
                        //Amarillo #FFF5D9
                        case decimal n when (n >= (promedio * 1.05M) && n < (promedio * 1.10M)):
                            color = "#FFF5D9";
                            break;
                        //Rojo  #FFC9C9
                        case decimal n when (n > (promedio * 1.10M)):
                            color = "#FFC9C9";
                            break;
                    }
                }
                else
                {
                    switch (valor)
                    {
                        //Blanco    #F2F2F2
                        case decimal n when ((n >= (promedio * 0.95M) && n <= (promedio * 1.05M))):
                            color = "#F2F2F2";
                            break;
                        //Verde #DEEDD3
                        case decimal n when (n > (promedio * 1.05M)):
                            color = "#DEEDD3";
                            break;
                        //Amarillo #FFF5D9
                        case decimal n when (n >= (promedio * 0.90M) && n < (promedio * 0.95M)):
                            color = "#FFF5D9";
                            break;
                        //Rojo  #FFC9C9
                        case decimal n when (n < (promedio * 0.90M)):
                            color = "#FFC9C9";
                            break;
                    }
                }
            }
            return color;
        }

        private void LlenarListas(DateTime fechaIni, DateTime fechaFin, int horaCorte)
        {
            string mensaje = "";
            controlador1.DatosTeoricos(ran_id, horaCorte, fechaIni, fechaFin, out listaProduccion, out listaSecas, out listaReto, out listaDestet1, out listaDestete2,
                out listaVaquillas, ref mensaje);
        }
        #endregion
    }
}
