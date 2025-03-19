using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReportePeriodo.Model.Impl;
using gth = LibreriaGTH.Consumos.Modelo;
using gthEntity = LibreriaGTH.Consumos.Entidad;
using ReportePeriodo.Entity;
using ReportePeriodo.Data;
using System.Data;
using System.Data.Common;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TreeView;

namespace ReportePeriodo.Model
{
    public class ModelReporte
    {
        #region variables de conexion de base de datos
        string _sqlConnection;
        string _oleDbConnection;
        string _fdbConnection;
        #endregion

        gth.ConsumoModel consumoModel;

        #region variables LibreriaGTH
        List<gthEntity.IndicadorReporteTeorico> indicadorProduccion;
        List<gthEntity.IndicadorReporteTeorico> indicadorSecas;
        List<gthEntity.IndicadorReporteTeorico> indicadorReto;
        List<gthEntity.IndicadorReporteTeorico> indicadorCrecimiento;
        List<gthEntity.IndicadorReporteTeorico> indicadorDesarrollo;
        List<gthEntity.IndicadorReporteTeorico> indicadorVaquillas;

        gthEntity.IndicadorReportePeriodoTyp promedioProduccion;
        gthEntity.IndicadorReportePeriodoTyp promedioSecas;
        gthEntity.IndicadorReportePeriodoTyp promedioReto;
        gthEntity.IndicadorReportePeriodoTyp promedioJaulas;
        gthEntity.IndicadorReportePeriodoTyp promedioCrecimiento;
        gthEntity.IndicadorReportePeriodoTyp promedioDesarrollo;
        gthEntity.IndicadorReportePeriodoTyp promedioVaquillas;

        gthEntity.IndicadorReportePeriodoTyp promedioAnteriorProduccion;
        gthEntity.IndicadorReportePeriodoTyp promedioAnteriorSecas;
        gthEntity.IndicadorReportePeriodoTyp promedioAnteriorReto;
        gthEntity.IndicadorReportePeriodoTyp promedioAnteriorJaulas;
        gthEntity.IndicadorReportePeriodoTyp promedioAnteriorCrecimiento;
        gthEntity.IndicadorReportePeriodoTyp promedioAnteriorDesarrollo;
        gthEntity.IndicadorReportePeriodoTyp promedioAnteriorVaquillas;

        gthEntity.EtapaTyp etapaProduccion = new gthEntity.EtapaTyp() { Id = 1, Clave = "10", Nombre = "Produccion" };
        gthEntity.EtapaTyp etapaSecas = new gthEntity.EtapaTyp() { Id = 2, Clave = "21", Nombre = "Secas" };
        gthEntity.EtapaTyp etapaReto = new gthEntity.EtapaTyp() { Id = 3, Clave = "22", Nombre = "Reto" };
        gthEntity.EtapaTyp etapaJaulas = new gthEntity.EtapaTyp() { Id = 4, Clave = "31", Nombre = "Jaulas" };
        gthEntity.EtapaTyp etapaCrecimiento = new gthEntity.EtapaTyp() { Id = 5, Clave = "32", Nombre = "Crecimiento" };
        gthEntity.EtapaTyp etapaDesarrollo = new gthEntity.EtapaTyp() { Id = 6, Clave = "33", Nombre = "Desarrollo" };
        gthEntity.EtapaTyp etapaVaquillas = new gthEntity.EtapaTyp() { Id = 7, Clave = "34", Nombre = "Vaquillas Preñadas" };
        #endregion

        #region variables
        List<DatosProduccion> _datosDeProduccion;
        List<NivelCribas> _datosNivelDeCribas;
        List<Inventario> _datosInventario;
        List<CalostroYOrdeña> _datosCalostro;
        List<Mproduc> _datosMproduc;

        List<InventarioAfiXDia> _datosInventarioAFI;
        List<Mproduc2> _datosMproduc2;

        List<DatosVacas> _datosJaulasFuera;
        List<DatosVacas> _datosJaulasVivas;
        List<DatosVacas> _datosJaulasMuertas;
        List<DatosVacas> _datosDesteteVivas;
        List<DatosVacas> _datosDesteteMuertas;
        List<DatosVacas> _datosVaquillasMuertas;
        List<DatosVacas> _datosVaquillasUrgencia;
        List<DatosVacas> _datosVaquillasDelgadas;
        List<DatosVacas> _datosVaquillasRegulares;
        List<DatosVacas> _datosVaquillasGordas;
        List<DatosVacas> _datosVaquillasOtros;
        List<DatosVacas> _datosVacasMuertas;
        List<DatosVacas> _datosVacasUrgencia;
        List<DatosVacas> _datosVacasDelgadas;
        List<DatosVacas> _datosVacasRegulares;
        List<DatosVacas> _datosVacasGordas;
        List<DatosVacas> _datosVacasOtros;
        List<DatosVacas> _datosPartosVaquillasND;
        List<DatosVacas> _datosPartosVaquillasAbortos;
        List<DatosVacas> _datosPartosVaquillasHembra;
        List<DatosVacas> _datosPartosVaquillasMacho;
        List<DatosVacas> _datosPartosVacasND;
        List<DatosVacas> _datosPartosVacasAbortos;
        List<DatosVacas> _datosPartosVacasHembra;
        List<DatosVacas> _datosPartosVacasMacho;
        List<DatosVacas> _datosPartosSinCria;
        List<DatosVacas> _datosPartosVacas;
        List<DatosVacas> _datosPartosVaquillas;
        List<DatosVacas> _datosMuertasDia;
        List<DatosVacas> _datosMuertasNoche;

        List<DatosVacas> _datosUbre_MA;
        List<DatosVacas> _datosUbre_SL;
        List<DatosVacas> _datosMetabolicos_FL;
        List<DatosVacas> _datosMetabolicos_CET;
        List<DatosVacas> _datosLocomotores_BE;
        List<DatosVacas> _datosLocomotores_TRA;
        List<DatosVacas> _datosLocomotores_GA;
        List<DatosVacas> _datosDigestivos_AC;
        List<DatosVacas> _datosDigestivos_ES;
        List<DatosVacas> _datosDigestivos_DI;
        List<DatosVacas> _datosDigestivos_TI;
        List<DatosVacas> _datosReproductivos_RE;
        List<DatosVacas> _datosReproductivos_ME;
        List<DatosVacas> _datosReproductivos_PIO;
        List<DatosVacas> _datosReproductivos_QUI;
        List<DatosVacas> _datosReproductivos_CS;
        List<DatosVacas> _datosRespiratorios_Neu;
        List<DatosVacas> _datosBecerras_Neu;
        List<DatosVacas> _datosBecerras_Fie;
        List<DatosVacas> _datosBecerras_Di;
        List<DatosVacas> _datosBecerras_Conj;
        List<ValInventario> _datosValInventarios;
        List<DatosVacas> _datosAbortosVacas;
        List<DatosVacas> _datosAbortosVaquillas;

        PromedioHoja1 _promedioHoja1;
        PromedioHoja1 _promedioHoja1AñoAnt;
        CalostroYOrdeña _promedioCalostro;
        CalostroYOrdeña _promedioCalostroAñoAnt;
        DatosProduccion _promedioDatosProd;
        DatosProduccion _promedioDatosProdAñoAnt;

        InventarioAfiXDia _promedioInventarioAFI;
        InventarioAfiXDia _promedioInventarioAFIAñoAnt;
        decimal _precioLeche;
        decimal _precioLecheAñoAnt;
        decimal _lecheProducida;
        decimal _lecheProducidaAñoAnt;

        Hoja1 _totalHoja1;
        Hoja1 _totalHoja1AñoAnt;

        Hoja2 _totalHoja2;
        Hoja2 _totalHoja2AñoAnt;

        Hoja3 _totalHoja3;
        Hoja3 _totalHoja3AñoAnt;

        Hoja4 _totalHoja4;
        Hoja4 _totalHoja4AñoAnt;
        #endregion

        #region Reportes
        List<Hoja1> _hoja1List;
        List<Hoja2> _hoja2List;
        List<Hoja3> _hoja3List;
        List<Hoja4> _hoja4List;
        #endregion



        public ModelReporte(string oleDbConnection, string sqlConnection, string fbdConnection)
        {
            this._oleDbConnection = oleDbConnection;
            this._sqlConnection = sqlConnection;
            this._fdbConnection = fbdConnection;
            this.consumoModel = new gth.ConsumoModel(sqlConnection);

            indicadorProduccion = new List<gthEntity.IndicadorReporteTeorico>();
            indicadorSecas = new List<gthEntity.IndicadorReporteTeorico>();
            indicadorReto = new List<gthEntity.IndicadorReporteTeorico>();
            indicadorCrecimiento = new List<gthEntity.IndicadorReporteTeorico>();
            indicadorDesarrollo = new List<gthEntity.IndicadorReporteTeorico>();
            indicadorVaquillas = new List<gthEntity.IndicadorReporteTeorico>();

            promedioProduccion = new gthEntity.IndicadorReportePeriodoTyp();
            promedioSecas = new gthEntity.IndicadorReportePeriodoTyp();
            promedioReto = new gthEntity.IndicadorReportePeriodoTyp();
            promedioJaulas = new gthEntity.IndicadorReportePeriodoTyp();
            promedioCrecimiento = new gthEntity.IndicadorReportePeriodoTyp();
            promedioDesarrollo = new gthEntity.IndicadorReportePeriodoTyp();
            promedioVaquillas = new gthEntity.IndicadorReportePeriodoTyp();

            promedioAnteriorProduccion = new gthEntity.IndicadorReportePeriodoTyp();
            promedioAnteriorSecas = new gthEntity.IndicadorReportePeriodoTyp();
            promedioAnteriorReto = new gthEntity.IndicadorReportePeriodoTyp();
            promedioAnteriorJaulas = new gthEntity.IndicadorReportePeriodoTyp();
            promedioAnteriorCrecimiento = new gthEntity.IndicadorReportePeriodoTyp();
            promedioAnteriorDesarrollo = new gthEntity.IndicadorReportePeriodoTyp();
            promedioAnteriorVaquillas = new gthEntity.IndicadorReportePeriodoTyp();

            _datosDeProduccion = new List<DatosProduccion>();
            _datosNivelDeCribas = new List<NivelCribas>();
            _datosInventario = new List<Inventario>();
            _datosCalostro = new List<CalostroYOrdeña>();
            _datosMproduc = new List<Mproduc>();

            _datosInventarioAFI = new List<InventarioAfiXDia>();
            _datosMproduc2 = new List<Mproduc2>();

            _datosJaulasFuera = new List<DatosVacas>();
            _datosJaulasVivas = new List<DatosVacas>();
            _datosJaulasMuertas = new List<DatosVacas>();
            _datosDesteteVivas = new List<DatosVacas>();
            _datosDesteteMuertas = new List<DatosVacas>();
            _datosVaquillasMuertas = new List<DatosVacas>();
            _datosVaquillasUrgencia = new List<DatosVacas>();
            _datosVaquillasDelgadas = new List<DatosVacas>();
            _datosVaquillasRegulares = new List<DatosVacas>();
            _datosVaquillasGordas = new List<DatosVacas>();
            _datosVaquillasOtros = new List<DatosVacas>();
            _datosVacasMuertas = new List<DatosVacas>();
            _datosVacasUrgencia = new List<DatosVacas>();
            _datosVacasDelgadas = new List<DatosVacas>();
            _datosVacasRegulares = new List<DatosVacas>();
            _datosVacasGordas = new List<DatosVacas>();
            _datosVacasOtros = new List<DatosVacas>();
            _datosPartosVaquillasND = new List<DatosVacas>();
            _datosPartosVaquillasAbortos = new List<DatosVacas>();
            _datosPartosVaquillasHembra = new List<DatosVacas>();
            _datosPartosVaquillasMacho = new List<DatosVacas>();
            _datosPartosVacasND = new List<DatosVacas>();
            _datosPartosVacasAbortos = new List<DatosVacas>();
            _datosPartosVacasHembra = new List<DatosVacas>();
            _datosPartosVacasMacho = new List<DatosVacas>();
            _datosPartosSinCria = new List<DatosVacas>();
            _datosPartosVacas = new List<DatosVacas>();
            _datosPartosVaquillas = new List<DatosVacas>();
            _datosMuertasDia = new List<DatosVacas>();
            _datosMuertasNoche = new List<DatosVacas>();

            _datosUbre_MA = new List<DatosVacas>();
            _datosUbre_SL = new List<DatosVacas>();
            _datosMetabolicos_FL = new List<DatosVacas>();
            _datosMetabolicos_CET = new List<DatosVacas>();
            _datosLocomotores_BE = new List<DatosVacas>();
            _datosLocomotores_TRA = new List<DatosVacas>();
            _datosLocomotores_GA = new List<DatosVacas>();
            _datosDigestivos_AC = new List<DatosVacas>();
            _datosDigestivos_ES = new List<DatosVacas>();
            _datosDigestivos_DI = new List<DatosVacas>();
            _datosDigestivos_TI = new List<DatosVacas>();
            _datosReproductivos_RE = new List<DatosVacas>();
            _datosReproductivos_ME = new List<DatosVacas>();
            _datosReproductivos_PIO = new List<DatosVacas>();
            _datosReproductivos_QUI = new List<DatosVacas>();
            _datosReproductivos_CS = new List<DatosVacas>();
            _datosRespiratorios_Neu = new List<DatosVacas>();
            _datosBecerras_Neu = new List<DatosVacas>();
            _datosBecerras_Fie = new List<DatosVacas>();
            _datosBecerras_Di = new List<DatosVacas>();
            _datosBecerras_Conj = new List<DatosVacas>();
            _datosValInventarios = new List<ValInventario>();
            _datosAbortosVacas = new List<DatosVacas>();
            _datosAbortosVaquillas = new List<DatosVacas>();

            _promedioHoja1 = new PromedioHoja1();
            _promedioHoja1AñoAnt = new PromedioHoja1();
            _promedioCalostro = new CalostroYOrdeña();
            _promedioCalostroAñoAnt = new CalostroYOrdeña();

            _promedioDatosProd = new DatosProduccion();
            _promedioDatosProdAñoAnt = new DatosProduccion();

            _promedioInventarioAFI = new InventarioAfiXDia();
            _promedioInventarioAFIAñoAnt = new InventarioAfiXDia();
            _precioLeche = 0;
            _precioLecheAñoAnt = 0;
            _lecheProducida = 0;
            _lecheProducidaAñoAnt = 0;

            _totalHoja3 = new Hoja3();
            _totalHoja3AñoAnt = new Hoja3();

            _totalHoja4 = new Hoja4();
            _totalHoja4AñoAnt = new Hoja4();

            _hoja1List = new List<Hoja1>();
            _hoja2List = new List<Hoja2>();
            _hoja3List = new List<Hoja3>();
            _hoja4List = new List<Hoja4>();
        }

        #region Propiedades
        public string SqlConnection
        {
            get { return _sqlConnection; }
            set { _sqlConnection = value; }
        }

        public string OleDbConnection
        {
            get { return _oleDbConnection; }
            set { _oleDbConnection = value; }
        }

        public string FdbConnection
        {
            get { return _fdbConnection; }
            set { _fdbConnection = value; }
        }

        public Hoja1 TotalHoja1
        {
            get
            {
                if (_totalHoja1 == null)
                    _totalHoja1 = new Hoja1() { Dia = "PROM" };

                return _totalHoja1;
            }
        }

        public Hoja1 TotalHoja1AñoAnt
        {
            get
            {
                if (_totalHoja1AñoAnt == null)
                    _totalHoja1AñoAnt = new Hoja1() { Dia = "AÑO ANT" };

                return _totalHoja1AñoAnt;
            }
        }

        public Hoja2 TotalHoja2
        {
            get
            {
                if (_totalHoja2 == null)
                    _totalHoja2 = new Hoja2() { Dia = "PROM" };

                return _totalHoja2;
            }
        }

        public Hoja2 TotalHoja2AñoAnt
        {
            get
            {
                if (_totalHoja2AñoAnt == null)
                    _totalHoja2AñoAnt = new Hoja2() { Dia = "AÑO ANT" };

                return _totalHoja2AñoAnt;
            }
        }
        
        
        #endregion

        #region metodos
        public gthEntity.CentroTyp GetCentro()
        {
            return consumoModel.GetCentro();
        }

        public Rancho GetRancho()
        {
            Rancho response = new Rancho();
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT rl.RanchoLocal  AS Id
                                        , r.Nombre      AS Nombre
                                        , r.Empresa     AS EmpresaId
                                        , r.NOIDREAL    AS NoIdReal
                                        , r.VERLITROS   AS VerLitros
                                        , r.pesadores   AS Pesadores
                                 FROM rancholocal rl
                                 LEFT JOIN ranchos r on rl.rancholocal = r.Clave";

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                response = LibreriaGTH.Utils.ConvertToList<Rancho>(db.EjecutarConsultaTabla()).FirstOrDefault();
            }
            catch (DbException ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "GetRancho", $"Error in {nameof(GetRancho)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        public DateTime GetFechaMaxima()
        {
            DateTime response = DateTime.Today;
            int juliana = LibreriaGTH.Utils.ConvertToJulian(response);
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT MAX(FECHA) FROM VALINVENTARIO";

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                DataTable dt = db.EjecutarConsultaTabla();

                if (dt.Rows.Count > 0)
                    juliana = dt.Rows[0][0] != DBNull.Value ? Convert.ToInt32(dt.Rows[0][0]) : LibreriaGTH.Utils.ConvertToJulian(response);

                response = ConvertJulianaToGregoriana(juliana);
            }
            catch (DbException ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "FechaMaxima", $"Error in {nameof(GetFechaMaxima)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        public DateTime GetFechaMinima()
        {
            DateTime response = DateTime.Today;
            int juliana = LibreriaGTH.Utils.ConvertToJulian(response);
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT MIN(FECHA) FROM VALINVENTARIO";

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                DataTable dt = db.EjecutarConsultaTabla();

                if (dt.Rows.Count > 0)
                    juliana = dt.Rows[0][0] != DBNull.Value ? Convert.ToInt32(dt.Rows[0][0]) : LibreriaGTH.Utils.ConvertToJulian(response);

                response = ConvertJulianaToGregoriana(juliana);
            }
            catch (DbException ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "FechaMinima", $"Error in {nameof(GetFechaMinima)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        public void CargarDatos(gthEntity.CentroTyp centro, Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {

            #region hoja 1 y 2 
            var datosproduccionTask = new Task(() => { DatosDeProduccion(fechaInicio, fechaFin); });
            var nivelCribasTask = new Task(() => { NivelCribas(fechaInicio, fechaFin); });
            var inventarioTask = new Task(() => { Inventario(fechaInicio, fechaFin); });
            var calostroTask = new Task(() => { CalostroOrdeño(fechaInicio, fechaFin); });
            var mproducTask = new Task(() => { Mproduc(rancho, fechaInicio, fechaFin); });
            var inventarioAFITask = new Task(() => { InventarioAFIxDIA(fechaInicio, fechaFin); });
            var mproducTask2 = new Task(() => { Mediaproduccion2(fechaInicio, fechaFin); });
            #endregion

            #region hoja 3
            var jaulasFueraTask = new Task(() => { JaulasFuera(fechaInicio, fechaFin); });
            var jaulasVivasTask = new Task(() => { JaulasVivas(fechaInicio, fechaFin); });
            var jaulasMuertasTask = new Task(() => { JaulasMuertas(fechaInicio, fechaFin); });
            var desteteVivasTask = new Task(() => { DesteteVivas(fechaInicio, fechaFin); });
            var desteteMuertasTask = new Task(() => { DesteteMuertas(fechaInicio, fechaFin); });
            var vaquillasMuertasTask = new Task(() => { VaquillasMuertas(fechaInicio, fechaFin); });
            var vaquillasUrgenciaTask = new Task(() => { VaquillasUrgencia(fechaInicio, fechaFin); });
            var vaquillasDelgadasTask = new Task(() => { VaquillasDelgadas(fechaInicio, fechaFin); });
            var vaquillasRegularesTask = new Task(() => { VaquillasRegulares(fechaInicio, fechaFin); });
            var vaquillasGordasTask = new Task(() => { VaquillasGordas(fechaInicio, fechaFin); });
            var vaquillasOtrosTask = new Task(() => { VaquillasOtros(fechaInicio, fechaFin); });

            var vacasMuertasTask = new Task(() => { VacasMuertas(fechaInicio, fechaFin); });
            var vacasUrgenciaTask = new Task(() => { VacasUrgencia(fechaInicio, fechaFin); });
            var vacasDelgadasTask = new Task(() => { VacasDelgadas(fechaInicio, fechaFin); });
            var vacasRegularesTask = new Task(() => { VacasRegulares(fechaInicio, fechaFin); });
            var vacasGordasTask = new Task(() => { VacasGordas(fechaInicio, fechaFin); });
            var vacasOtrosTask = new Task(() => { VacasOtros(fechaInicio, fechaFin); });

            var partosVaquillasNDTask = new Task(() => { PartosVaquillasND(fechaInicio, fechaFin); });
            var partosVaquillasAbortosTask = new Task(() => { PartosVaquillasAbortos(fechaInicio, fechaFin); });
            var partosVaquillasHembraTask = new Task(() => { PartosVaquillasHembra(fechaInicio, fechaFin); });
            var partosVaquillasMachoTask = new Task(() => { PartosVaquillasMacho(fechaInicio, fechaFin); });

            var partosVacasNDTask = new Task(() => { PartosVacasND(fechaInicio, fechaFin); });
            var partosVacasAbortosTask = new Task(() => { PartosVacasAbortos(fechaInicio, fechaFin); });
            var partosVacasHembraTask = new Task(() => { PartosVacasHembra(fechaInicio, fechaFin); });
            var partosVacasMachoTask = new Task(() => { PartosVacasMacho(fechaInicio, fechaFin); });
            var partosSinCriaTask = new Task(() => { PartosSinCria(fechaInicio, fechaFin); });
            var partosVacasTask = new Task(() => { PartosVacas(fechaInicio, fechaFin); });
            var partosVaquillasTask = new Task(() => { PartosVaquillas(fechaInicio, fechaFin); });
            var muertasDiaTask = new Task(() => { MuertasDia(fechaInicio, fechaFin); });
            var muertasNocheTask = new Task(() => { MuertasNoche(fechaInicio, fechaFin); });
            #endregion

            #region hoja 4
            var ubreTask = new Task(() => { Ubre(fechaInicio, fechaFin); });
            var metabolicosTask = new Task(() => { Metabolicos(fechaInicio, fechaFin); });
            var locomotoresTask = new Task(() => { Locomotores(fechaInicio, fechaFin); });
            var digestivosTask = new Task(() => { Digestivos(fechaInicio, fechaFin); });
            var reproductivosTask = new Task(() => { Reproductivos(fechaInicio, fechaFin); });
            var respiratoriosTask = new Task(() => { Respiratorios(fechaInicio, fechaFin); });
            var becerrasTask = new Task(() => { Becerras(fechaInicio, fechaFin); });
            var valInventariosTask = new Task(() => { ValidacionInventarios(fechaInicio, fechaFin); });
            var abortosTask = new Task(() => { Abortos(fechaInicio, fechaFin); });
            #endregion

            #region promedios
            var promedioHoja1Task = new Task(() => { PromedioHoja1(rancho, fechaInicio, fechaFin); });
            var promedioCalostroTask = new Task(() => { PromedioCalostro(fechaInicio, fechaFin); });
            var totalDatosDeProduccionTask = new Task(() => { TotalDatosDeProduccion(fechaInicio, fechaFin); });
            var pecioLecheTask = new Task(() => { PrecioLeche(centro.CentroClave, fechaInicio, fechaFin); });
            var lecheProducidaTask = new Task(() => { LecheProducida(fechaInicio, fechaFin); });
            var totalesHoja3Task = new Task(() => { TotalesHoja3(rancho, fechaInicio, fechaFin); });
            var totalesHoja4Task = new Task(() => { TotalesHoja4(rancho, fechaInicio, fechaFin); });
            #endregion


            #region inicio de tareas
            datosproduccionTask.Start();
            nivelCribasTask.Start();
            inventarioTask.Start();
            calostroTask.Start();
            mproducTask.Start();
            inventarioAFITask.Start();
            mproducTask2.Start();

            jaulasFueraTask.Start();
            jaulasVivasTask.Start();
            jaulasMuertasTask.Start();
            desteteVivasTask.Start();
            desteteMuertasTask.Start();
            vaquillasMuertasTask.Start();
            vaquillasUrgenciaTask.Start();
            vaquillasDelgadasTask.Start();
            vaquillasRegularesTask.Start();
            vaquillasGordasTask.Start();
            vaquillasOtrosTask.Start();

            vacasMuertasTask.Start();
            vacasUrgenciaTask.Start();
            vacasDelgadasTask.Start();
            vacasRegularesTask.Start();
            vacasGordasTask.Start();
            vacasOtrosTask.Start();

            partosVaquillasNDTask.Start();
            partosVaquillasAbortosTask.Start();
            partosVaquillasHembraTask.Start();
            partosVaquillasMachoTask.Start();

            partosVacasNDTask.Start();
            partosVacasAbortosTask.Start();
            partosVacasHembraTask.Start();
            partosVacasMachoTask.Start();
            partosSinCriaTask.Start();
            partosVacasTask.Start();
            partosVaquillasTask.Start();
            muertasDiaTask.Start();
            muertasNocheTask.Start();

            ubreTask.Start();
            metabolicosTask.Start();
            locomotoresTask.Start();
            digestivosTask.Start();
            reproductivosTask.Start();
            respiratoriosTask.Start();
            becerrasTask.Start();
            valInventariosTask.Start();
            abortosTask.Start();

            promedioHoja1Task.Start();
            promedioCalostroTask.Start();
            totalDatosDeProduccionTask.Start();
            pecioLecheTask.Start();
            lecheProducidaTask.Start();
            totalesHoja3Task.Start();
            totalesHoja4Task.Start();
            #endregion


            var info = new List<Task>
            {
                datosproduccionTask,
                nivelCribasTask,
                inventarioTask,
                calostroTask,
                mproducTask,
                inventarioAFITask,
                mproducTask2,
                jaulasFueraTask,
                jaulasVivasTask,
                jaulasMuertasTask,
                desteteVivasTask,
                desteteMuertasTask,
                vaquillasMuertasTask,
                vaquillasUrgenciaTask,
                vaquillasDelgadasTask,
                vaquillasRegularesTask,
                vaquillasGordasTask,
                vaquillasOtrosTask,
                vacasMuertasTask,
                vacasUrgenciaTask,
                vacasDelgadasTask,
                vacasRegularesTask,
                vacasGordasTask,
                vacasOtrosTask,
                partosVaquillasNDTask,
                partosVaquillasAbortosTask,
                partosVaquillasHembraTask,
                partosVaquillasMachoTask,
                partosVacasNDTask,
                partosVacasAbortosTask,
                partosVacasHembraTask,
                partosVacasMachoTask,
                partosSinCriaTask,
                partosVacasTask,
                partosVaquillasTask,
                muertasDiaTask,
                muertasNocheTask,
                ubreTask,
                metabolicosTask,
                locomotoresTask,
                digestivosTask,
                reproductivosTask,
                respiratoriosTask,
                becerrasTask,
                valInventariosTask,
                abortosTask,
                promedioHoja1Task,
                promedioCalostroTask,
                totalDatosDeProduccionTask,
                pecioLecheTask,
                lecheProducidaTask,
                totalesHoja3Task,
                totalesHoja4Task
            };

            Task.WaitAll(info.ToArray());

        }

        public void CargarDatosTeoricos(gthEntity.CentroTyp centro, DateTime fechaInicio, DateTime fechaFin)
        {
            var produccionTask = new Task(() => { IndicadoresTeoricoProduccion(centro, fechaInicio, fechaFin); });
            var secasTask = new Task(() => { IndicadoresTeoricoSecas(centro, fechaInicio, fechaFin); });
            var retoTask = new Task(() => { IndicadoresTeoricoReto(centro, fechaInicio, fechaFin); });
            var crecimientoTask = new Task(() => { IndicadoresTeoricoCrecimiento(centro, fechaInicio, fechaFin); });
            var desarrolloTask = new Task(() => { IndicadoresTeoricoDesarrollo(centro, fechaInicio, fechaFin); });
            var vaquillasTask = new Task(() => { IndicadoresTeoricoVaquillas(centro, fechaInicio, fechaFin); });

            produccionTask.Start();
            secasTask.Start();
            retoTask.Start();
            crecimientoTask.Start();
            desarrolloTask.Start();
            vaquillasTask.Start();

            var info = new List<Task>
            {
                produccionTask,
                secasTask,
                retoTask,
                crecimientoTask,
                desarrolloTask,
                vaquillasTask
            };

            Task.WaitAll(info.ToArray());
        }

        public void CargarDatosAlimentacion(gthEntity.CentroTyp centro, DateTime fechaInicio, DateTime fechaFin)
        {
            var produccionTask = new Task(() => { IndicadorProduccion(centro, fechaInicio, fechaFin); });
            var secasTask = new Task(() => { IndicadorSecas(centro, fechaInicio, fechaFin); });
            var retoTask = new Task(() => { IndicadorReto(centro, fechaInicio, fechaFin); });
            var jaulasTask = new Task(() => { IndicadorJaulas(centro, fechaInicio, fechaFin); });
            var crecimientoTask = new Task(() => { IndicadorCrecimiento(centro, fechaInicio, fechaFin); });
            var desarrolloTask = new Task(() => { IndicadorDesarrollo(centro, fechaInicio, fechaFin); });
            var vaquillasTask = new Task(() => { IndicadorVaquillas(centro, fechaInicio, fechaFin); });

            produccionTask.Start();
            secasTask.Start();
            retoTask.Start();
            jaulasTask.Start();
            crecimientoTask.Start();
            desarrolloTask.Start();
            vaquillasTask.Start();

            var info = new List<Task>
            {
                produccionTask,
                secasTask,
                retoTask,
                jaulasTask,
                crecimientoTask,
                desarrolloTask,
                vaquillasTask
            };

            Task.WaitAll(info.ToArray());
        }
        #endregion

        #region Llenar Hojas
        private List<Hoja1> GetReporteHoja1(Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            List<Hoja1> response = new List<Hoja1>();
            List<DateTime> fechas = GetFechas(fechaFin);

            try
            {
                foreach (DateTime fecha in fechas)
                {
                    Hoja1 hoja = new Hoja1();
                    hoja.Dia = fecha.Day.ToString();

                    DatosProduccion busquedaDatosProduccion = (from x in _datosDeProduccion where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    Mproduc busquedaMProduc = (from x in _datosMproduc where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    NivelCribas busquedaCribas = (from x in _datosNivelDeCribas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    CalostroYOrdeña busquedaCalostro = (from x in _datosCalostro where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    Inventario busquedaInventario = (from x in _datosInventario where x.Fecha == fecha select x).ToList().FirstOrDefault();

                    hoja.Ordeño = busquedaMProduc != null ? busquedaMProduc.Vacas_Ordeño : null;
                    hoja.Secas = busquedaMProduc != null ? busquedaMProduc.Vacas_Secas : null;
                    hoja.Hato = busquedaMProduc != null ? busquedaMProduc.Vacas_Hato : null;
                    hoja.Porcentaje_Lact = busquedaMProduc != null ? busquedaMProduc.Lactosa : null;
                    hoja.Porcentaje_Prot = busquedaMProduc != null ? busquedaMProduc.Proteina : null;
                    hoja.Urea = busquedaDatosProduccion != null ? busquedaDatosProduccion.Urea : null;
                    hoja.Porcentaje_Grasa = busquedaDatosProduccion != null ? busquedaDatosProduccion.Grasa : null;
                    hoja.CCS = busquedaDatosProduccion != null ? busquedaDatosProduccion.CCS : null;
                    hoja.CTD = busquedaDatosProduccion != null ? busquedaDatosProduccion.CTD : null;

                    hoja.Leche = busquedaMProduc != null ? busquedaMProduc.Leche_Producida : null;
                    hoja.Antib = busquedaMProduc != null ? busquedaMProduc.Antib : null;
                    hoja.Media = busquedaMProduc != null ? busquedaMProduc.Media : null;
                    hoja.Total = busquedaMProduc != null ? busquedaMProduc.Total_Producido : null;
                    hoja.DEL = busquedaInventario != null ? busquedaInventario.DelOrd : null;
                    hoja.ANT = busquedaMProduc != null ? busquedaMProduc.Vaca_Antib : null;

                    hoja.EA = busquedaMProduc != null ? busquedaMProduc.EA : null;
                    hoja.ILCA = busquedaMProduc != null ? busquedaMProduc.ILCA_Produccion : null;
                    hoja.IC = busquedaMProduc != null ? busquedaMProduc.IC_Produccion : null;
                    hoja.Costo_Litro = busquedaMProduc != null ? busquedaMProduc.Media != null && busquedaMProduc.Costo != null ? busquedaMProduc.Media != 0 && busquedaMProduc.Costo != 0 ? busquedaMProduc.Costo / busquedaMProduc.Media : null : null : null;
                    hoja.MH = busquedaMProduc != null ? busquedaMProduc.MH : null;
                    hoja.MS = busquedaMProduc != null ? busquedaMProduc.MS : null;
                    hoja.Porcentaje_MS = hoja.MH != null && hoja.MS != null ? hoja.MH != 0 ? hoja.MS / hoja.MH : null : null;
                    hoja.SA = busquedaMProduc != null ? busquedaMProduc.SA : null;
                    hoja.MSS = busquedaMProduc != null ? busquedaMProduc.MSS : null;
                    hoja.EAS = busquedaMProduc != null ? busquedaMProduc.EAS : null;
                    hoja.Porcentaje_Sob = hoja.SA != null && hoja.MH != null ? hoja.SA != 0 && hoja.MH != 0 ? hoja.SA / hoja.MH * 100 : null : null;
                    hoja.Costo_Prod = busquedaMProduc != null ? busquedaMProduc.Costo : null;
                    hoja.Costo_MS = hoja.Costo_Prod != null && hoja.MS != null ? hoja.Costo_Prod != 0 && hoja.MS != 0 ? hoja.Costo_Prod / hoja.MS : null : null;

                    hoja.Cribas_N1 = busquedaCribas != null ? busquedaCribas.Nivel1 : null;
                    hoja.Cribas_N2 = busquedaCribas != null ? busquedaCribas.Nivel2 : null;
                    hoja.Cribas_N3 = busquedaCribas != null ? busquedaCribas.Nivel3 : null;
                    hoja.Cribas_N4 = busquedaCribas != null ? busquedaCribas.Nivel4 : null;

                    hoja.NoID_Ses1 = busquedaMProduc != null ? busquedaMProduc.Noid_Ses1 : null;
                    hoja.NoID_Ses2 = busquedaMProduc != null ? busquedaMProduc.Noid_Ses2 : null;
                    hoja.NoID_Ses3 = busquedaMProduc != null ? busquedaMProduc.Noid_Ses3 : null;

                    hoja.SES1 = busquedaCalostro != null ? rancho.VerLitros ? busquedaCalostro.Litros_Sesion1 : busquedaCalostro.Horario_Sesion1 : null;
                    hoja.SES2 = busquedaCalostro != null ? rancho.VerLitros ? busquedaCalostro.Litros_Sesion2 : busquedaCalostro.Horario_Sesion2 : null;
                    hoja.SES3 = busquedaCalostro != null ? rancho.VerLitros ? busquedaCalostro.Litros_Sesion3 : busquedaCalostro.Horario_Sesion3 : null;

                    hoja.Porcentaje_Revueltas = busquedaCalostro != null ? busquedaCalostro.Porcentaje_Revueltas : null;

                    response.Add(hoja);
                }
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "ReporteHoja1", $"Error in {nameof(GetReporteHoja1)}", ex.Message, ex.Message, ex.StackTrace, guid);
            }

            return response;

        }

        public List<Hoja2> GetReporteHoja2(Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            List<Hoja2> response = new List<Hoja2>();
            List<DateTime> fechas = GetFechas(fechaFin);

            try
            {
                foreach (DateTime fecha in fechas)
                {
                    InventarioAfiXDia busqudaInventarioAFI = (from x in _datosInventarioAFI where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    Mproduc2 busquedaMproduc = (from x in _datosMproduc2 where x.Fecha == fecha select x).ToList().FirstOrDefault();

                    Hoja2 item = new Hoja2();
                    item.Dia = fecha.Day.ToString();

                    #region Crianza
                    item.Crianza_Inventario = busqudaInventarioAFI != null ? busqudaInventarioAFI.Crianza : 0;
                    #endregion 

                    #region Jaulas 
                    item.Jaulas_Inventario = busqudaInventarioAFI != null ? busqudaInventarioAFI.Jaulas : 0;
                    item.Jaulas_Costo = busquedaMproduc != null ? busquedaMproduc.Jaulas_Costo : 0;
                    #endregion

                    #region Crecimiento 
                    item.Crecimiento_Inventario = busqudaInventarioAFI != null ? busqudaInventarioAFI.Crecimiento : 0;
                    item.Crecimiento_MH = busquedaMproduc != null ? busquedaMproduc.Crecimiento_MH : 0;
                    item.Crecimiento_Costo = busquedaMproduc != null ? busquedaMproduc.Crecimiento_Costo : 0;
                    item.Crecimiento_PorcentajeMS = busquedaMproduc != null ? busquedaMproduc.Crecimiento_Porcentaje_MS : 0;
                    item.Crecimiento_MS = busquedaMproduc != null ? busquedaMproduc.Crecimiento_MS : 0;
                    item.Crecimiento_CostoMS = busquedaMproduc != null ? busquedaMproduc.Crecimiento_Costo_MS : 0;
                    #endregion

                    #region Desarrollo 
                    item.Desarrollo_Inventario = busqudaInventarioAFI != null ? busqudaInventarioAFI.Desarrollo : 0;
                    item.Desarrollo_MH = busquedaMproduc != null ? busquedaMproduc.Desarrollo_MH : 0;
                    item.Desarrollo_Costo = busquedaMproduc != null ? busquedaMproduc.Desarrollo_Costo : 0;
                    item.Desarrollo_PorcentajeMS = busquedaMproduc != null ? busquedaMproduc.Desarrollo_Porcentaje_MS : 0;
                    item.Desarrollo_MS = busquedaMproduc != null ? busquedaMproduc.Desarrollo_MS : 0;
                    item.Desarrollo_CostoMS = busquedaMproduc != null ? busquedaMproduc.Desarrollo_Costo_MS : 0;
                    #endregion

                    #region Vaquillas 
                    item.Vaquillas_Inventario = busqudaInventarioAFI != null ? busqudaInventarioAFI.Vaquillas : 0;
                    item.Vaquillas_MH = busquedaMproduc != null ? busquedaMproduc.Vaquillas_MH : 0;
                    item.Vaquillas_Costo = busquedaMproduc != null ? busquedaMproduc.Vaquillas_Costo : 0;
                    item.Vaquillas_PorcentajeMS = busquedaMproduc != null ? busquedaMproduc.Vaquillas_Porcentaje_MS : 0;
                    item.Vaquillas_MS = busquedaMproduc != null ? busquedaMproduc.Vaquillas_MS : 0;
                    item.Vaquillas_CostoMS = busquedaMproduc != null ? busquedaMproduc.Vaquillas_Costo_MS : 0;
                    #endregion

                    #region Secas 
                    item.Secas_Inventario = busqudaInventarioAFI != null ? busqudaInventarioAFI.Secas : 0;
                    item.Secas_MH = busquedaMproduc != null ? busquedaMproduc.Secas_MH : 0;
                    item.Secas_PorcentajeMS = busquedaMproduc != null ? busquedaMproduc.Secas_Porcentaje_MS : 0;
                    item.Secas_MS = busquedaMproduc != null ? busquedaMproduc.Secas_MS : 0;
                    item.Secas_SA = busquedaMproduc != null ? busquedaMproduc.Secas_SA : 0;
                    item.Secas_Mss = busquedaMproduc != null ? busquedaMproduc.Secas_MSS : 0;
                    item.Secas_PorcentajeSob = busquedaMproduc != null ? busquedaMproduc.Secas_Porcentaje_SA : 0;
                    item.Secas_Costo = busquedaMproduc != null ? busquedaMproduc.Secas_Costo : 0;
                    item.Secas_CostoMS = busquedaMproduc != null ? busquedaMproduc.Secas_Costo_MS : 0;
                    #endregion

                    #region Reto 
                    item.Reto_Inventario = busqudaInventarioAFI != null ? busqudaInventarioAFI.Reto : 0;
                    item.Reto_MH = busquedaMproduc != null ? busquedaMproduc.Reto_MH : 0;
                    item.Reto_PorcentajeMS = busquedaMproduc != null ? busquedaMproduc.Reto_Porcentaje_MS : 0;
                    item.Reto_MS = busquedaMproduc != null ? busquedaMproduc.Reto_MS : 0;
                    item.Reto_SA = busquedaMproduc != null ? busquedaMproduc.Reto_SA : 0;
                    item.Reto_Mss = busquedaMproduc != null ? busquedaMproduc.Reto_MSS : 0;
                    item.Reto_PorcentajeSob = busquedaMproduc != null ? busquedaMproduc.Reto_Porcentaje_SA : 0;
                    item.Reto_Costo = busquedaMproduc != null ? busquedaMproduc.Reto_Costo : 0;
                    item.Reto_CostoMS = busquedaMproduc != null ? busquedaMproduc.Reto_Costo_MS : 0;
                    #endregion

                    #region Utilidad 
                    item.Inventario_Total = busqudaInventarioAFI != null ? busqudaInventarioAFI.InventarioTotal : 0;
                    item.IngresoxAnimal = busquedaMproduc != null ? item.Inventario_Total > 0 ? (busquedaMproduc.LecheProd * _precioLeche) / item.Inventario_Total : 0 : 0;

                    decimal? auxJaulas = item.Jaulas_Costo * item.Jaulas_Inventario;
                    decimal? auxCrecimiento = item.Crecimiento_Costo * item.Crecimiento_Inventario;
                    decimal? auxDesarrollo = item.Desarrollo_Costo * item.Desarrollo_Inventario;
                    decimal? auxVaquillas = item.Vaquillas_Costo * item.Vaquillas_Inventario;
                    decimal? auxSecas = item.Secas_Costo * item.Secas_Inventario;
                    decimal? auxReto = item.Reto_Costo * item.Reto_Inventario;
                    decimal? auxOrdeño = busquedaMproduc != null && busqudaInventarioAFI != null ? busqudaInventarioAFI.Ordeño * busquedaMproduc.Ordeño_Costo : 0;
                    decimal? auxCosto = auxJaulas + auxCrecimiento + auxDesarrollo + auxVaquillas + auxSecas + auxReto + auxOrdeño;
                    item.CostoxAnimal = item.Inventario_Total > 0 ? auxCosto / item.Inventario_Total : 0;
                    item.UtilidadxAnimal = item.IngresoxAnimal - item.CostoxAnimal;
                    item.Porcentaje_CostoxAnimal = item.IngresoxAnimal > 0 ? item.CostoxAnimal / item.IngresoxAnimal : 0;
                    item.Porcentaje_UtilidadxAnimal = item.IngresoxAnimal > 0 ? item.UtilidadxAnimal / item.IngresoxAnimal : 0;

                    #endregion


                    response.Add(item);
                }
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "ReporteHoja2", $"Error in {nameof(GetReporteHoja2)}", ex.Message, ex.Message, ex.StackTrace, guid);
            }

            return response;
        }

        #endregion

        #region promedios

        public void CargarPromedios(Rancho rancho)
        {
            var promedioHoja1Task = new Task(() => { PromediosHoja1(rancho); });
            var promedioHoja2Task = new Task(() => { PromediosHoja2(); });
            var promedioHoja3Task = new Task(() => { PromediosHoja3(); });
            var promedioHoja4Task = new Task(() => { PromediosHoja4(); });


            promedioHoja1Task.Start();
            promedioHoja2Task.Start();
            promedioHoja3Task.Start();
            promedioHoja4Task.Start();

            var infoTasks = new List<Task>
            {
                promedioHoja1Task,
                promedioHoja2Task,
                promedioHoja3Task,
                promedioHoja4Task
            };


            Task.WaitAll(infoTasks.ToArray());
        }


        #region Hoja 1
        public void CargarPromediosHoja1(Rancho rancho)
        {
            _totalHoja1 = GetPromedioHoja1(rancho, "PROM", promedioProduccion, _promedioHoja1, _promedioCalostro, _promedioDatosProd);
            _totalHoja1AñoAnt = GetPromedioHoja1(rancho, "AÑO ANT", promedioAnteriorProduccion, _promedioHoja1AñoAnt, _promedioCalostroAñoAnt, _promedioDatosProdAñoAnt);
        }

        private Hoja1 GetPromedioHoja1(Rancho rancho, string dia, gthEntity.IndicadorReportePeriodoTyp indicadorOrdeño, PromedioHoja1 promedioH1, CalostroYOrdeña promedioCalostro, DatosProduccion promedioDatos)
        {
            Hoja1 response = new Hoja1() { Dia = dia };

            try
            {
                response.Ordeño = promedioH1.Ordeño;
                response.Secas = promedioH1.Secas;
                response.Hato = promedioH1.Hato;
                response.Porcentaje_Lact = promedioH1.Porcentaje_Lact;
                response.Porcentaje_Prot = promedioH1.Porcentaje_Prot;
                response.Urea = promedioDatos.Urea;
                response.Porcentaje_Grasa = promedioDatos.Grasa;
                response.CCS = promedioDatos.CCS;
                response.CTD = promedioDatos.CTD;
                response.DEC = promedioDatos.Grasa;

                response.Leche = promedioH1.Leche;
                response.Antib = promedioH1.Antib;
                response.Media = indicadorOrdeño.Media;
                response.Total = promedioH1.Total;
                response.DEL = promedioH1.DEL;
                response.ANT = promedioH1.Ant;

                response.EA = indicadorOrdeño.EficienciaAlimentacion;
                response.ILCA = indicadorOrdeño.ILCAProducccion;
                response.IC = indicadorOrdeño.ICProduccion;
                response.Costo_Litro = indicadorOrdeño.CostoLitro;
                response.MH = indicadorOrdeño.MateriaHumeda;
                response.Porcentaje_MS = (indicadorOrdeño.PorcentajeMS / 100);
                response.MS = indicadorOrdeño.MateriaSeca;
                response.SA = indicadorOrdeño.SobranteAnimal;
                response.MSS = indicadorOrdeño.MateriaSecaSob;
                response.EAS = indicadorOrdeño.EficienciaAlimentacionSob;
                response.Porcentaje_Sob = indicadorOrdeño.MateriaHumeda > 0 ? Convert.ToDecimal(indicadorOrdeño.SobranteAnimal / indicadorOrdeño.MateriaHumeda * 100) : 0;
                response.Costo_Prod = indicadorOrdeño.CostoRacion;
                response.Costo_MS = indicadorOrdeño.CostoMateriaSeca;

                response.Cribas_N1 = promedioH1.Cribas_N1;
                response.Cribas_N2 = promedioH1.Cribas_N2;
                response.Cribas_N3 = promedioH1.Cribas_N3;
                response.Cribas_N4 = promedioH1.Cribas_N4;

                response.NoID_Ses1 = promedioH1.Sesion1;
                response.NoID_Ses2 = promedioH1.Sesion2;
                response.NoID_Ses3 = promedioH1.Sesion3;

                response.SES1 = rancho.VerLitros ? promedioCalostro.Litros_Sesion1 : promedioCalostro.Horario_Sesion1;
                response.SES2 = rancho.VerLitros ? promedioCalostro.Litros_Sesion2 : promedioCalostro.Horario_Sesion2;
                response.SES3 = rancho.VerLitros ? promedioCalostro.Litros_Sesion3 : promedioCalostro.Horario_Sesion3;

                response.Porcentaje_Revueltas = promedioCalostro.Porcentaje_Revueltas;
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "PromedioHoja1", $"Error in {nameof(GetPromedioHoja1)}", ex.Message, ex.Message, ex.StackTrace, guid);
            }

            return response;
        }
        #endregion

        #region Hoja 2
        public void CargarPromediosHoja2()
        {
            _totalHoja2 = GetPromedioHoja2("PROM", promedioProduccion, promedioSecas, promedioReto, promedioJaulas, promedioCrecimiento, promedioDesarrollo, promedioVaquillas, _promedioInventarioAFI, _precioLeche, _lecheProducida);
            _totalHoja2AñoAnt = GetPromedioHoja2("AÑO ANT", promedioAnteriorProduccion, promedioAnteriorSecas, promedioAnteriorReto, promedioAnteriorJaulas, promedioAnteriorCrecimiento, promedioAnteriorDesarrollo, promedioAnteriorVaquillas, _promedioInventarioAFIAñoAnt, _precioLecheAñoAnt, _lecheProducidaAñoAnt);        
        }

        private Hoja2 GetPromedioHoja2(string dia, gthEntity.IndicadorReportePeriodoTyp indicadorOrdeño, gthEntity.IndicadorReportePeriodoTyp indicadorSecas, gthEntity.IndicadorReportePeriodoTyp indicadorReto,
            gthEntity.IndicadorReportePeriodoTyp indicadorJaulas, gthEntity.IndicadorReportePeriodoTyp indicadorCrecimiento, gthEntity.IndicadorReportePeriodoTyp indicadorDesarrollo,
            gthEntity.IndicadorReportePeriodoTyp indicadorVaquillas, InventarioAfiXDia inventario, decimal precioLeche, decimal lecheProducida)
        {
            Hoja2 response = new Hoja2() { Dia = dia };

            try
            {
                response.Crianza_Inventario = inventario.Crianza;

                response.Jaulas_Inventario = indicadorJaulas.NumeroAnimal;
                response.Jaulas_Costo = indicadorJaulas.CostoRacion;

                response.Crecimiento_Inventario = indicadorCrecimiento.NumeroAnimal;
                response.Crecimiento_MH = indicadorCrecimiento.MateriaHumeda;
                response.Crecimiento_MS = indicadorCrecimiento.MateriaSeca;
                response.Crecimiento_PorcentajeMS = indicadorCrecimiento.PorcentajeMS;
                response.Crecimiento_Costo = indicadorCrecimiento.CostoRacion;
                response.Crecimiento_CostoMS = indicadorCrecimiento.CostoMateriaSeca;

                response.Desarrollo_Inventario = indicadorDesarrollo.NumeroAnimal;
                response.Desarrollo_MH = indicadorDesarrollo.MateriaHumeda;
                response.Desarrollo_MS = indicadorDesarrollo.MateriaSeca;
                response.Desarrollo_PorcentajeMS = indicadorDesarrollo.PorcentajeMS;
                response.Desarrollo_Costo = indicadorDesarrollo.CostoRacion;
                response.Desarrollo_CostoMS = indicadorDesarrollo.CostoMateriaSeca;

                response.Vaquillas_Inventario = indicadorVaquillas.NumeroAnimal;
                response.Vaquillas_MH = indicadorVaquillas.MateriaHumeda;
                response.Vaquillas_MS = indicadorVaquillas.MateriaSeca;
                response.Vaquillas_PorcentajeMS = indicadorVaquillas.PorcentajeMS;
                response.Vaquillas_Costo = indicadorVaquillas.CostoRacion;
                response.Vaquillas_CostoMS = indicadorVaquillas.CostoMateriaSeca;

                response.Secas_Inventario = indicadorSecas.NumeroAnimal;
                response.Secas_MH = indicadorSecas.MateriaHumeda;
                response.Secas_PorcentajeMS = indicadorSecas.PorcentajeMS;
                response.Secas_MS = indicadorSecas.MateriaSeca;
                response.Secas_SA = indicadorSecas.SobranteAnimal;
                response.Secas_Mss = indicadorSecas.MateriaSecaSob;
                response.Secas_PorcentajeSob = indicadorSecas.MateriaHumeda > 0 ? (indicadorSecas.SobranteAnimal / indicadorSecas.MateriaHumeda * 100) : 0;
                response.Secas_Costo = indicadorSecas.CostoRacion;
                response.Secas_CostoMS = indicadorSecas.CostoMateriaSeca;

                response.Reto_Inventario = indicadorReto.NumeroAnimal;
                response.Reto_MH = indicadorReto.MateriaHumeda;
                response.Reto_PorcentajeMS = indicadorReto.PorcentajeMS;
                response.Reto_MS = indicadorReto.MateriaSeca;
                response.Reto_SA = indicadorReto.SobranteAnimal;
                response.Reto_Mss = indicadorReto.MateriaSecaSob;
                response.Reto_PorcentajeSob = indicadorReto.MateriaHumeda > 0 ? (indicadorReto.SobranteAnimal / indicadorReto.MateriaHumeda * 100) : 0;
                response.Reto_Costo = indicadorReto.CostoRacion;
                response.Reto_CostoMS = indicadorReto.CostoMateriaSeca;

                response.Inventario_Total = inventario.InventarioTotal;

                decimal ixA = 0;
                try { ixA = response.Inventario_Total > 0 ? Convert.ToDecimal((lecheProducida * precioLeche) / response.Inventario_Total) : 0; }
                catch { }

                response.IngresoxAnimal = Math.Round(ixA, 5);

                decimal? auxJaulas = response.Jaulas_Costo * response.Jaulas_Inventario;
                decimal? auxCrecimiento = response.Crecimiento_Costo * response.Crecimiento_Inventario;
                decimal? auxDesarrollo = response.Desarrollo_Costo * response.Desarrollo_Inventario;
                decimal? auxVaquillas = response.Vaquillas_Costo * response.Vaquillas_Inventario;
                decimal? auxSecas = response.Secas_Costo * response.Secas_Inventario;
                decimal? auxReto = response.Reto_Costo * response.Reto_Inventario;
                decimal? auxOrdeño = Convert.ToDecimal(indicadorOrdeño.NumeroAnimal * indicadorOrdeño.CostoRacion);
                decimal? auxCosto = (auxJaulas ?? 0) + (auxCrecimiento ?? 0) + (auxDesarrollo ?? 0) + (auxVaquillas ?? 0) + (auxSecas ?? 0) + (auxReto ?? 0) + (auxOrdeño ?? 0);

                decimal cxa = 0;
                try { cxa = response.Inventario_Total > 0 ? Convert.ToDecimal(auxCosto / response.Inventario_Total) : 0; }
                catch { }

                response.CostoxAnimal = Math.Round(cxa, 5);
                response.UtilidadxAnimal = response.IngresoxAnimal - response.CostoxAnimal;
                response.Porcentaje_CostoxAnimal = response.IngresoxAnimal > 0 ? response.CostoxAnimal / response.IngresoxAnimal : 0;
                response.Porcentaje_UtilidadxAnimal = response.IngresoxAnimal > 0 ? response.UtilidadxAnimal / response.IngresoxAnimal : 0;
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "PromedioHoja2", $"Error in {nameof(GetPromedioHoja2)}", ex.Message, ex.Message, ex.StackTrace, guid);
            }

            return response;
        }


        #endregion

        #region Hoja 3

        #endregion

        #region Hoja 4

        #endregion

        #endregion


        #region metodos privados
        private DateTime ConvertJulianaToGregoriana(int request)
        {
            return new DateTime(1900, 1, 1).AddDays(request).AddDays(-2);
        }

        private List<DateTime> GetFechas(DateTime fechaFin)
        {
            List<DateTime> response = new List<DateTime>();
            DateTime fechaInicio = new DateTime(fechaFin.Year, fechaFin.Month, 1);

            while (fechaInicio <= fechaFin)
            {
                response.Add(fechaInicio);
                fechaInicio = fechaInicio.AddDays(1);
            }

            return response;
        }
        #endregion

        #region Tareas Asincronas

        #region Hoja 1 y 2
        private async Task<bool> DatosDeProduccion(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosDeProduccion = GetDatosDeProduccion(fechaInicio, fechaFin);
            Console.WriteLine("Proceso {0} Ejecutado.", "DatosDeProduccion");
            return true;
        }

        private async Task<bool> NivelCribas(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosNivelDeCribas = GetNivelCribas(fechaInicio, fechaFin);
            Console.WriteLine("Proceso {0} Ejecutado.", "NivelCribas");
            return true;
        }

        private async Task<bool> Inventario(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosInventario = GetInventario(fechaInicio, fechaFin);
            Console.WriteLine("Proceso {0} Ejecutado.", "Inventario");
            return true;
        }

        private async Task<bool> CalostroOrdeño(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosCalostro = GetCalostroOrdeño(fechaInicio, fechaFin);
            Console.WriteLine("Proceso {0} Ejecutado.", "CalostroOrdeño");
            return true;
        }

        private async Task<bool> Mproduc(Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            _datosMproduc = rancho.NoIdReal ? GetMediaProduccionNoIdReal(fechaInicio, fechaFin) : GetMediaProduccionNoId(fechaInicio, fechaFin);
            Console.WriteLine("Proceso {0} Ejecutado.", "Mproduc");
            return true;
        }

        private async Task<bool> InventarioAFIxDIA(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosInventarioAFI = GetInventariosAFI(fechaInicio, fechaFin);
            Console.WriteLine("Proceso {0} Ejecutado.", "InventarioAFIxDIA");
            return true;
        }

        private async Task<bool> Mediaproduccion2(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosMproduc2 = GetMediaProduccion2(fechaInicio, fechaFin);
            Console.WriteLine("Proceso {0} Ejecutado.", "Mediaproduccion2");
            return true;
        }

        private async Task<bool> PromedioHoja1(Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            _promedioHoja1 = rancho.NoIdReal ? GetPromedioHoja1NoIdReal(fechaInicio, fechaFin) : GetPromedioHoja1NoId(fechaInicio, fechaFin);
            _promedioHoja1AñoAnt = rancho.NoIdReal ? GetPromedioHoja1NoIdReal(fechaInicio.AddYears(-1), fechaFin.AddYears(-1)) : GetPromedioHoja1NoId(fechaInicio.AddYears(-1), fechaFin.AddYears(-1));
            Console.WriteLine("Proceso {0} Ejecutado.", "PromedioHoja1NoId");
            return true;
        }

        private async Task<bool> PromedioCalostro(DateTime fechaInicio, DateTime fechaFin)
        {
            _promedioCalostro = GetPromedioCalostro(fechaInicio, fechaFin);
            _promedioCalostroAñoAnt = GetPromedioCalostro(fechaInicio.AddYears(-1), fechaFin.AddYears(-1));
            Console.WriteLine("Proceso {0} Ejecutado.", "PromedioCalostro");
            return true;
        }

        private async Task<bool> TotalDatosDeProduccion(DateTime fechaInicio, DateTime fechaFin)
        {
            _promedioDatosProd = GetTotalDatosDeProduccion(fechaInicio, fechaFin);
            _promedioDatosProdAñoAnt = GetTotalDatosDeProduccion(fechaInicio.AddYears(-1), fechaFin.AddYears(-1));
            Console.WriteLine("Proceso {0} Ejecutado.", "TotalDatosDeProduccion");
            return true;
        }

        private async Task<bool> PrecioLeche(int centroId, DateTime fechaInicio, DateTime fechaFin)
        {
            bool esPrecioLecheFacturado = EsPrecioLecheFacturado(fechaFin);
            _precioLeche = esPrecioLecheFacturado ? GetPrecioLecheFacturado(fechaInicio, fechaFin) : GetPrecioLecheCalculado(centroId, fechaInicio, fechaFin);
            _precioLecheAñoAnt = GetPrecioLecheFacturado(fechaInicio.AddYears(-1), fechaFin.AddYears(-1));
            Console.WriteLine("Proceso {0} Ejecutado.", "PrecioLeche");
            return true;
        }

        private async Task<bool> LecheProducida(DateTime fechaInicio, DateTime fechaFin)
        {
            _lecheProducida = GetLecheProducida(fechaInicio, fechaFin);
            _lecheProducidaAñoAnt = GetLecheProducida(fechaInicio.AddYears(-1), fechaFin.AddYears(-1));
            Console.WriteLine("Proceso {0} Ejecutado.", "LecheProducida");
            return true;
        }

        private async Task<bool> PromediosHoja1(Rancho rancho)
        {
            CargarPromediosHoja1(rancho);
            Console.WriteLine("Proceso {0} Ejecutado.", "PromediosHoja1");
            return true;
        }

        private async Task<bool> PromediosHoja2()
        {
            CargarPromediosHoja2();
            Console.WriteLine("Proceso {0} Ejecutado.", "PromediosHoja2");
            return true;
        }

        private async Task<bool> PromediosHoja3()
        {

            Console.WriteLine("Proceso {0} Ejecutado.", "PromediosHoja3");
            return true;
        }

        private async Task<bool> PromediosHoja4()
        {

            Console.WriteLine("Proceso {0} Ejecutado.", "PromediosHoja4");
            return true;
        }

        #endregion

        #region Hoja 3
        private async Task<bool> JaulasFuera(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosJaulasFuera = GetDatosDesecho(fechaInicio, fechaFin, @"AND LACTANCIA = 0 AND EDOVAC = 15");
            Console.WriteLine("Proceso {0} Ejecutado.", "JaulasFuera");
            return true;
        }

        private async Task<bool> JaulasVivas(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosJaulasVivas = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo <> 1 AND destino <> 2 AND (edovac = 1 OR edovac = 10)");
            Console.WriteLine("Proceso {0} Ejecutado.", "JaulasVivas");
            return true;
        }

        private async Task<bool> JaulasMuertas(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosJaulasMuertas = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 1 AND (edovac = 1 OR edovac = 10) ");
            Console.WriteLine("Proceso {0} Ejecutado.", "JaulasMuertas");
            return true;
        }

        private async Task<bool> DesteteVivas(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosDesteteVivas = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo <> 1 AND destino <> 2 AND (edovac = 2 OR edovac = 7) ");
            Console.WriteLine("Proceso {0} Ejecutado.", "DesteteVivas");
            return true;
        }

        private async Task<bool> DesteteMuertas(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosDesteteMuertas = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 1 AND (edovac = 2 OR edovac = 7) ");
            Console.WriteLine("Proceso {0} Ejecutado.", "DesteteMuertas");
            return true;
        }

        private async Task<bool> VaquillasMuertas(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosVaquillasMuertas = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 1 AND destino = 1  AND (edovac = 3 OR edovac = 8 OR edovac = 9) ");
            Console.WriteLine("Proceso {0} Ejecutado.", "VaquillasMuertas");
            return true;
        }

        private async Task<bool> VaquillasUrgencia(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosVaquillasUrgencia = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 2 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ");
            Console.WriteLine("Proceso {0} Ejecutado.", "VaquillasUrgencia");
            return true;
        }

        private async Task<bool> VaquillasDelgadas(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosVaquillasDelgadas = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 3 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ");
            Console.WriteLine("Proceso {0} Ejecutado.", "VaquillasDelgadas");
            return true;
        }

        private async Task<bool> VaquillasRegulares(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosVaquillasRegulares = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 4 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ");
            Console.WriteLine("Proceso {0} Ejecutado.", "VaquillasRegulares");
            return true;
        }

        private async Task<bool> VaquillasGordas(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosVaquillasGordas = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 5 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ");
            Console.WriteLine("Proceso {0} Ejecutado.", "VaquillasGordas");
            return true;
        }

        private async Task<bool> VaquillasOtros(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosVaquillasOtros = GetDatosDesecho(fechaInicio, fechaFin, @"  AND vacvaq = 2  AND motivo > 5 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ");
            Console.WriteLine("Proceso {0} Ejecutado.", "VaquillasOtros");
            return true;
        }

        private async Task<bool> VacasMuertas(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosVacasMuertas = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo = 1 AND destino = 1 AND (edovac > 3) ");
            Console.WriteLine("Proceso {0} Ejecutado.", "VacasMuertas");
            return true;
        }

        private async Task<bool> VacasUrgencia(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosVacasUrgencia = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo = 2 AND destino = 1 AND (edovac > 3) ");
            Console.WriteLine("Proceso {0} Ejecutado.", "VacasUrgencia");
            return true;
        }

        private async Task<bool> VacasDelgadas(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosVacasDelgadas = GetDatosDesecho(fechaInicio, fechaFin, @"  AND vacvaq = 1  AND motivo = 3 AND destino = 1 AND (edovac > 3) ");
            Console.WriteLine("Proceso {0} Ejecutado.", "VacasDelgadas");
            return true;
        }

        private async Task<bool> VacasRegulares(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosVacasRegulares = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo = 4 AND destino = 1 AND (edovac > 3) ");
            Console.WriteLine("Proceso {0} Ejecutado.", "VacasRegulares");
            return true;
        }

        private async Task<bool> VacasGordas(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosVacasGordas = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo = 5 AND destino = 1 AND (edovac > 3) ");
            Console.WriteLine("Proceso {0} Ejecutado.", "VacasGordas");
            return true;
        }

        private async Task<bool> VacasOtros(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosVacasOtros = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo > 5 AND destino = 1 AND (edovac > 3) ");
            Console.WriteLine("Proceso {0} Ejecutado.", "VacasOtros");
            return true;
        }

        private async Task<bool> PartosVaquillasND(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosPartosVaquillasND = GetDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2  AND (tipopar = 1 or tipopar = 3) ");
            Console.WriteLine("Proceso {0} Ejecutado.", "PartosVaquillasND");
            return true;
        }

        private async Task<bool> PartosVaquillasAbortos(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosPartosVaquillasAbortos = GetDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2  AND (tipopar = 2) ");
            Console.WriteLine("Proceso {0} Ejecutado.", "PartosVaquillasAbortos");
            return true;
        }

        private async Task<bool> PartosVaquillasHembra(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosPartosVaquillasHembra = GetDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2  AND criaviva = 1  AND criasexo = 1 ");
            Console.WriteLine("Proceso {0} Ejecutado.", "PartosVaquillasHembra");
            return true;
        }

        private async Task<bool> PartosVaquillasMacho(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosPartosVaquillasMacho = GetDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2  AND criaviva = 1  AND criasexo = 2 ");
            Console.WriteLine("Proceso {0} Ejecutado.", "PartosVaquillasMacho");
            return true;
        }

        private async Task<bool> PartosVacasND(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosPartosVacasND = GetDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 1  AND (tipopar = 1 or tipopar = 3) ");
            Console.WriteLine("Proceso {0} Ejecutado.", "PartosVacasND");
            return true;
        }

        private async Task<bool> PartosVacasAbortos(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosPartosVacasAbortos = GetDatosNacimiento(fechaInicio, fechaFin, @"  AND vacvaq = 1  AND (tipopar = 2) ");
            Console.WriteLine("Proceso {0} Ejecutado.", "PartosVacasAbortos");
            return true;
        }

        private async Task<bool> PartosVacasHembra(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosPartosVacasHembra = GetDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 1  AND criaviva = 1  AND criasexo = 1 ");
            Console.WriteLine("Proceso {0} Ejecutado.", "PartosVacasHembra");
            return true;
        }

        private async Task<bool> PartosVacasMacho(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosPartosVacasMacho = GetDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 1  AND criaviva = 1  AND criasexo = 2 ");
            Console.WriteLine("Proceso {0} Ejecutado.", "PartosVacasMacho");
            return true;
        }

        private async Task<bool> PartosSinCria(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosPartosSinCria = GetDatosNacimiento(fechaInicio, fechaFin, @" AND tipopar <> 2 AND criaviva = 2  AND criasexo = 3 ");
            Console.WriteLine("Proceso {0} Ejecutado.", "PartosSinCria");
            return true;
        }

        private async Task<bool> PartosVacas(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosPartosVacas = GetDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 1 ");
            Console.WriteLine("Proceso {0} Ejecutado.", "PartosVacas");
            return true;
        }

        private async Task<bool> PartosVaquillas(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosPartosVaquillas = GetDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2 ");
            Console.WriteLine("Proceso {0} Ejecutado.", "PartosVaquillas");
            return true;
        }

        private async Task<bool> MuertasDia(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosMuertasDia = GetDatosNacimientoMuertas(fechaInicio, fechaFin, @" AND criaviva = 2 AND (tipopar = 1 or tipopar = 3) AND criasexo < 3 AND dianoche = 1 ");
            Console.WriteLine("Proceso {0} Ejecutado.", "MuertasDia");
            return true;
        }

        private async Task<bool> MuertasNoche(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosMuertasNoche = GetDatosNacimientoMuertas(fechaInicio, fechaFin, @" AND criaviva = 2 AND (tipopar = 1 or tipopar = 3) AND criasexo < 3 AND dianoche = 2 ");
            Console.WriteLine("Proceso {0} Ejecutado.", "MuertasNoche");
            return true;
        }

        private async Task<bool> TotalesHoja3(Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            _totalHoja3 = GetTotalReporteHoja3(rancho, fechaInicio, fechaFin, "TOTAL");
            _totalHoja3AñoAnt = GetTotalReporteHoja3(rancho, fechaInicio.AddYears(-1), fechaFin.AddYears(-1), "AÑO ANT");
            Console.WriteLine("Proceso {0} Ejecutado.", "TotalesHoja3");
            return true;
        }

        private async Task<bool> TotalesHoja4(Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            _totalHoja4 = GetTotalReporteHoja4(rancho, fechaInicio, fechaFin, "TOTAL");
            _totalHoja4AñoAnt = GetTotalReporteHoja4(rancho, fechaInicio.AddYears(-1), fechaFin.AddYears(-1), "AÑO ANT");
            Console.WriteLine("Proceso {0} Ejecutado.", "TotalesHoja4");
            return true;
        }
        #endregion

        #region Hoja 4
        private async Task<bool> Ubre(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosUbre_MA = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 1 AND enfdetalle = 1 ");
            _datosUbre_SL = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 1 AND enfdetalle = 2 ");
            Console.WriteLine("Proceso {0} Ejecutado.", "Ubre");
            return true;
        }

        private async Task<bool> Metabolicos(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosMetabolicos_FL = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 5 AND enfdetalle = 1 ");
            _datosMetabolicos_CET = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 5 AND enfdetalle = 2 ");
            Console.WriteLine("Proceso {0} Ejecutado.", "Metabolicos");
            return true;
        }

        private async Task<bool> Locomotores(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosLocomotores_BE = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 4 AND enfdetalle = 1 ");
            _datosLocomotores_TRA = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 4 AND enfdetalle = 2 ");
            _datosLocomotores_GA = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 4 AND enfdetalle = 3 ");
            Console.WriteLine("Proceso {0} Ejecutado.", "Locomotores");
            return true;
        }

        private async Task<bool> Digestivos(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosDigestivos_AC = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 1 ");
            _datosDigestivos_ES = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 2 ");
            _datosDigestivos_DI = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 3 ");
            _datosDigestivos_TI = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 4 ");
            Console.WriteLine("Proceso {0} Ejecutado.", "Digestivos");
            return true;
        }

        private async Task<bool> Reproductivos(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosReproductivos_RE = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 1 ");
            _datosReproductivos_ME = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 2 ");
            _datosReproductivos_PIO = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 3 ");
            _datosReproductivos_QUI = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 4 ");
            _datosReproductivos_CS = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 5 ");
            Console.WriteLine("Proceso {0} Ejecutado.", "Reproductivos");
            return true;
        }

        private async Task<bool> Respiratorios(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosRespiratorios_Neu = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 7 AND enfdetalle = 1 ");
            Console.WriteLine("Proceso {0} Ejecutado.", "Respiratorios");
            return true;
        }

        private async Task<bool> Becerras(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosBecerras_Neu = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 2  AND enfermedad = 8 AND enfdetalle = 1 ");
            _datosBecerras_Fie = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 2  AND enfermedad = 8 AND enfdetalle = 2 ");
            _datosBecerras_Di = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 2  AND enfermedad = 8 AND enfdetalle = 3 ");
            _datosBecerras_Conj = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 2  AND enfermedad = 8 AND enfdetalle = 4 ");
            Console.WriteLine("Proceso {0} Ejecutado.", "Becerras");
            return true;
        }

        private async Task<bool> ValidacionInventarios(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosValInventarios = GetDatosValInventario(fechaInicio, fechaFin);
            Console.WriteLine("Proceso {0} Ejecutado.", "ValidacionInventarios");
            return true;
        }

        private async Task<bool> Abortos(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosAbortosVacas = GetDatosAbortos(fechaInicio, fechaFin, 1);
            _datosAbortosVaquillas = GetDatosAbortos(fechaInicio, fechaFin, 2);
            Console.WriteLine("Proceso {0} Ejecutado.", "Abortos");
            return true;
        }
        #endregion

        #region Datos Teoricos
        private async Task<bool> IndicadoresTeoricoProduccion(gthEntity.CentroTyp centro, DateTime fechaInicio, DateTime fechaFin)
        {
            List<gthEntity.CentroTyp> aux = new List<gthEntity.CentroTyp>();
            aux.Add(centro);
            indicadorProduccion = consumoModel.GetIndicadorReporteTeorico(aux, etapaProduccion, fechaInicio, fechaFin);
            Console.WriteLine("Proceso {0} Ejecutado.", "IndicadoresTeoricoProduccion");
            return true;
        }

        private async Task<bool> IndicadoresTeoricoSecas(gthEntity.CentroTyp centro, DateTime fechaInicio, DateTime fechaFin)
        {
            List<gthEntity.CentroTyp> aux = new List<gthEntity.CentroTyp>();
            aux.Add(centro);
            indicadorSecas = consumoModel.GetIndicadorReporteTeorico(aux, etapaSecas, fechaInicio, fechaFin);
            Console.WriteLine("Proceso {0} Ejecutado.", "IndicadoresTeoricoSecas");
            return true;
        }

        private async Task<bool> IndicadoresTeoricoReto(gthEntity.CentroTyp centro, DateTime fechaInicio, DateTime fechaFin)
        {
            List<gthEntity.CentroTyp> aux = new List<gthEntity.CentroTyp>();
            aux.Add(centro);
            indicadorReto = consumoModel.GetIndicadorReporteTeorico(aux, etapaReto, fechaInicio, fechaFin);
            Console.WriteLine("Proceso {0} Ejecutado.", "IndicadoresTeoricoReto");
            return true;
        }

        private async Task<bool> IndicadoresTeoricoCrecimiento(gthEntity.CentroTyp centro, DateTime fechaInicio, DateTime fechaFin)
        {
            List<gthEntity.CentroTyp> aux = new List<gthEntity.CentroTyp>();
            aux.Add(centro);
            indicadorCrecimiento = consumoModel.GetIndicadorReporteTeorico(aux, etapaCrecimiento, fechaInicio, fechaFin);
            Console.WriteLine("Proceso {0} Ejecutado.", "IndicadoresTeoricoCrecimiento");
            return true;
        }

        private async Task<bool> IndicadoresTeoricoDesarrollo(gthEntity.CentroTyp centro, DateTime fechaInicio, DateTime fechaFin)
        {
            List<gthEntity.CentroTyp> aux = new List<gthEntity.CentroTyp>();
            aux.Add(centro);
            indicadorDesarrollo = consumoModel.GetIndicadorReporteTeorico(aux, etapaDesarrollo, fechaInicio, fechaFin);
            Console.WriteLine("Proceso {0} Ejecutado.", "IndicadoresTeoricoDesarrollo");
            return true;
        }

        private async Task<bool> IndicadoresTeoricoVaquillas(gthEntity.CentroTyp centro, DateTime fechaInicio, DateTime fechaFin)
        {
            List<gthEntity.CentroTyp> aux = new List<gthEntity.CentroTyp>();
            aux.Add(centro);
            indicadorVaquillas = consumoModel.GetIndicadorReporteTeorico(aux, etapaVaquillas, fechaInicio, fechaFin);
            Console.WriteLine("Proceso {0} Ejecutado.", "IndicadoresTeoricoVaquillas");
            return true;
        }
        #endregion

        #region Datos de Alimentacion
        private async Task<bool> IndicadorProduccion(gthEntity.CentroTyp centro, DateTime fechaInicio, DateTime fechaFin)
        {
            try
            {
                List<gthEntity.CentroTyp> aux = new List<gthEntity.CentroTyp>();
                aux.Add(centro);

                promedioProduccion = consumoModel.GetIndicadorReportePeriodo(aux, etapaProduccion, fechaInicio, fechaFin);
                promedioAnteriorProduccion = consumoModel.GetIndicadorReportePeriodo(aux, etapaProduccion, fechaInicio.AddYears(-1), fechaFin.AddYears(-1));

                Console.WriteLine("Proceso {0} Ejecutado.", "IndicadorProduccion");
            }
            catch { }
            return true;
        }

        private async Task<bool> IndicadorSecas(gthEntity.CentroTyp centro, DateTime fechaInicio, DateTime fechaFin)
        {
            try
            {
                List<gthEntity.CentroTyp> aux = new List<gthEntity.CentroTyp>();
                aux.Add(centro);

                promedioSecas = consumoModel.GetIndicadorReportePeriodo(aux, etapaSecas, fechaInicio, fechaFin);
                promedioAnteriorSecas = consumoModel.GetIndicadorReportePeriodo(aux, etapaSecas, fechaInicio.AddYears(-1), fechaFin.AddYears(-1));

                Console.WriteLine("Proceso {0} Ejecutado.", "IndicadorSecas");
            }
            catch { }
            return true;
        }

        private async Task<bool> IndicadorReto(gthEntity.CentroTyp centro, DateTime fechaInicio, DateTime fechaFin)
        {
            try
            {
                List<gthEntity.CentroTyp> aux = new List<gthEntity.CentroTyp>();
                aux.Add(centro);

                promedioReto = consumoModel.GetIndicadorReportePeriodo(aux, etapaReto, fechaInicio, fechaFin);
                promedioAnteriorReto = consumoModel.GetIndicadorReportePeriodo(aux, etapaReto, fechaInicio.AddYears(-1), fechaFin.AddYears(-1));

                Console.WriteLine("Proceso {0} Ejecutado.", "IndicadorReto");
            }
            catch { }
            return true;
        }

        private async Task<bool> IndicadorJaulas(gthEntity.CentroTyp centro, DateTime fechaInicio, DateTime fechaFin)
        {
            try
            {
                List<gthEntity.CentroTyp> aux = new List<gthEntity.CentroTyp>();
                aux.Add(centro);

                promedioJaulas = consumoModel.GetIndicadorReportePeriodo(aux, etapaJaulas, fechaInicio, fechaFin);
                promedioAnteriorJaulas = consumoModel.GetIndicadorReportePeriodo(aux, etapaJaulas, fechaInicio.AddYears(-1), fechaFin.AddYears(-1));

                Console.WriteLine("Proceso {0} Ejecutado.", "IndicadorJaulas");
            }
            catch { }
            return true;
        }

        private async Task<bool> IndicadorCrecimiento(gthEntity.CentroTyp centro, DateTime fechaInicio, DateTime fechaFin)
        {
            try
            {
                List<gthEntity.CentroTyp> aux = new List<gthEntity.CentroTyp>();
                aux.Add(centro);

                promedioCrecimiento = consumoModel.GetIndicadorReportePeriodo(aux, etapaCrecimiento, fechaInicio, fechaFin);
                promedioAnteriorCrecimiento = consumoModel.GetIndicadorReportePeriodo(aux, etapaCrecimiento, fechaInicio.AddYears(-1), fechaFin.AddYears(-1));

                Console.WriteLine("Proceso {0} Ejecutado.", "IndicadorCrecimiento");
            }
            catch { }
            return true;
        }

        private async Task<bool> IndicadorDesarrollo(gthEntity.CentroTyp centro, DateTime fechaInicio, DateTime fechaFin)
        {
            try
            {
                List<gthEntity.CentroTyp> aux = new List<gthEntity.CentroTyp>();
                aux.Add(centro);

                promedioDesarrollo = consumoModel.GetIndicadorReportePeriodo(aux, etapaDesarrollo, fechaInicio, fechaFin);
                promedioAnteriorDesarrollo = consumoModel.GetIndicadorReportePeriodo(aux, etapaDesarrollo, fechaInicio.AddYears(-1), fechaFin.AddYears(-1));

                Console.WriteLine("Proceso {0} Ejecutado.", "IndicadorDesarrollo");
            }
            catch { }
            return true;
        }

        private async Task<bool> IndicadorVaquillas(gthEntity.CentroTyp centro, DateTime fechaInicio, DateTime fechaFin)
        {
            try
            {
                List<gthEntity.CentroTyp> aux = new List<gthEntity.CentroTyp>();
                aux.Add(centro);

                promedioVaquillas = consumoModel.GetIndicadorReportePeriodo(aux, etapaVaquillas, fechaInicio, fechaFin);
                promedioAnteriorVaquillas = consumoModel.GetIndicadorReportePeriodo(aux, etapaVaquillas, fechaInicio.AddYears(-1), fechaFin.AddYears(-1));

                Console.WriteLine("Proceso {0} Ejecutado.", "IndicadorVaquillas");
            }
            catch { }
            return true;
        }
        #endregion

        #endregion

        #region Hoja 1 y 2 
        private List<DatosProduccion> GetDatosDeProduccion(DateTime fechaInicio, DateTime fechaFin)
        {
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);
            List<DatosProduccion> response = new List<DatosProduccion>();

            try
            {
                string query = @"SELECT  CDATE(RESULTADOS.FECHA) AS FechaG
                                       ,AVG(RESULTADOS.PROTES)  AS Proteina1
                                       ,SUM(RESULTADOS.UREAS)   AS Urea1
                                       ,SUM(RESULTADOS.GRASAS)  AS Grasa1
                                       ,SUM(RESULTADOS.CCSS)    AS CCS1
                                       ,SUM(RESULTADOS.CTDS)    AS CTD1
                                FROM
                                (
	                                SELECT  LECHEXDIA.FECHA
	                                       ,VALORES.PROTEINA                                                                      AS PROTES
	                                       ,IIF(ISNULL(LECHEXDIA.LG),NULL,(VALORES.LITROSXTANQUE / LECHEXDIA.LG) * VALORES.GRASA) AS GRASAS
	                                       ,IIF(ISNULL(LECHEXDIA.LU),NULL,(VALORES.LITROSXTANQUE / LECHEXDIA.LU) * VALORES.UREA)  AS UREAS
	                                       ,IIF(ISNULL(LECHEXDIA.LCC),NULL,(VALORES.LITROSXTANQUE / LECHEXDIA.LCC) * VALORES.CCS) AS CCSS
	                                       ,IIF(ISNULL(LECHEXDIA.LCT),NULL,(VALORES.LITROSXTANQUE / LECHEXDIA.LCT) * VALORES.CTD) AS CTDS
	                                FROM
	                                (
		                                SELECT  LITROSxDIA.DIA             AS FECHA
		                                       ,SUM(LITROSxDIA.LITROGRASA) AS LG
		                                       ,SUM(LITROSxDIA.LITROUREA)  AS LU
		                                       ,SUM(LITROSxDIA.LITROCCS)   AS LCC
		                                       ,SUM(LITROSxDIA.LITROCTS)   AS LCT
		                                FROM
		                                (
			                                SELECT  FECHA                        AS DIA
			                                       ,GRASA
			                                       ,IIF(GRASA > 0,LITROSXTANQUE) AS LITROGRASA
			                                       ,IIF(UREA > 0,LITROSXTANQUE)  AS LITROUREA
			                                       ,IIF(CCS > 0,LITROSXTANQUE)   AS LITROCCS
			                                       ,IIF(CTD > 0,LITROSXTANQUE)   AS LITROCTS
			                                FROM dproduc
			                                WHERE FECHA BETWEEN @julianaI AND @julianaF
			                                ORDER BY FECHA 
		                                ) LITROSxDIA
		                                GROUP BY  LITROSxDIA.DIA
	                                ) LECHEXDIA
	                                LEFT JOIN
	                                (
		                                SELECT  FECHA
		                                       ,PROTEINA
		                                       ,LITROSXTANQUE
		                                       ,GRASA
		                                       ,UREA
		                                       ,CCS
		                                       ,CTD
		                                FROM dproduc
		                                WHERE FECHA BETWEEN @julianaI AND @julianaF
		                                ORDER BY FECHA 
	                                ) VALORES
	                                ON VALORES.FECHA = LECHEXDIA.FECHA
                                )RESULTADOS
                                GROUP BY  RESULTADOS.FECHA";

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@julianaI", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@julianaF", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new DatosProduccion()
                {
                    Fecha = x["FechaG"] != DBNull.Value ? Convert.ToDateTime(x["FechaG"]) : new DateTime(),
                    Proteina = x["Proteina1"] != DBNull.Value ? Convert.ToDecimal(x["Proteina1"]) : 0,
                    Urea = x["Urea1"] != DBNull.Value ? Convert.ToDecimal(x["Urea1"]) : 0,
                    Grasa = x["Grasa1"] != DBNull.Value ? Convert.ToDecimal(x["Grasa1"]) : 0,
                    CCS = x["CCS1"] != DBNull.Value ? Convert.ToDecimal(x["CCS1"]) : 0,
                    CTD = x["CTD1"] != DBNull.Value ? Convert.ToDecimal(x["CTD1"]) : 0
                }).ToList();
            }
            catch (DbException ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "GetDatosDeProduccion", $"Error in {nameof(GetDatosDeProduccion)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private List<NivelCribas> GetNivelCribas(DateTime fechaInicio, DateTime fechaFin)
        {
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);
            List<NivelCribas> response = new List<NivelCribas>();

            try
            {
                string query = @"SELECT  CDATE(FECHA) AS FechaG
                                       ,IIF(ISNULL(AVG(nivel1)),0,AVG(nivel1)) AS criba1
                                       ,IIF(ISNULL(AVG(nivel2)),0,AVG(nivel2)) AS criba2
                                       ,IIF(ISNULL(AVG(nivel3)),0,AVG(nivel3)) AS criba3
                                       ,IIF(ISNULL(AVG(nivel4)),0,AVG(nivel4)) AS criba4
                                FROM NIVELCRIBA
                                WHERE FECHA BETWEEN @julianaI AND @julianaF
                                GROUP BY  FECHA";

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@julianaI", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@julianaF", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new NivelCribas()
                {
                    Fecha = x["FechaG"] != DBNull.Value ? Convert.ToDateTime(x["FechaG"]) : new DateTime(),
                    Nivel1 = x["criba1"] != DBNull.Value ? Convert.ToDecimal(x["criba1"]) : 0,
                    Nivel2 = x["criba2"] != DBNull.Value ? Convert.ToDecimal(x["criba2"]) : 0,
                    Nivel3 = x["criba3"] != DBNull.Value ? Convert.ToDecimal(x["criba3"]) : 0,
                    Nivel4 = x["criba4"] != DBNull.Value ? Convert.ToDecimal(x["criba4"]) : 0

                }).ToList();

            }
            catch (DbException ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "GetNivelCribas", $"Error in {nameof(GetNivelCribas)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private List<Inventario> GetInventario(DateTime fechaInicio, DateTime fechaFin)
        {
            List<Inventario> response = new List<Inventario>();
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT  CDATE(Fecha) AS FechaG
                                        ,delord       AS Del
                                 FROM inventario
                                 WHERE fecha BETWEEN @fechaInicio AND @fechaFin";

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@fechaInicio", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new Inventario()
                {
                    Fecha = x["FechaG"] != DBNull.Value ? Convert.ToDateTime(x["FechaG"]) : new DateTime(),
                    DelOrd = x["Del"] != DBNull.Value ? Convert.ToDecimal(x["Del"]) : 0
                }).ToList();
            }
            catch (DbException ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "GetInventario", $"Error in {nameof(GetInventario)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private List<CalostroYOrdeña> GetCalostroOrdeño(DateTime fechaInicio, DateTime fechaFin)
        {
            List<CalostroYOrdeña> response = new List<CalostroYOrdeña>();
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT  CDATE(FECHA)                AS FechaG
										,IIF((ESPERADO_VQ + ESPERADO_VC) > 0,DIFERENCIA/(ESPERADO_VQ + ESPERADO_VC)*100,0)AS Porcentaje
										,DIFERENCIA
										,CALIDAD
										,HORARIO_DIFERENCIA1
										,HORARIO_DIFERENCIA2
										,HORARIO_DIFERENCIA3
										,CAPACIDADORDEÑO
										,PORCENREVUELTAS
										,LITROS1
										,LITROS2
										,LITROS3
                                        ,HORARIO_DIF_FININI1
                                        ,HORARIO_DIF_FININI2
                                        ,HORARIO_DIF_FININI3
									FROM CALOSTROYORDEÑODIA
									WHERE FECHA BETWEEN @fechaInicio AND @fechaFin";
                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@fechaInicio", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new CalostroYOrdeña()
                {
                    Fecha = x["FechaG"] != DBNull.Value ? Convert.ToDateTime(x["FechaG"]) : new DateTime(),
                    Porcentaje = x["Porcentaje"] != DBNull.Value ? Convert.ToDecimal(x["Porcentaje"]) : 0,
                    Diferencia = x["DIFERENCIA"] != DBNull.Value ? Convert.ToDecimal(x["DIFERENCIA"]) : 0,
                    Calidad = x["CALIDAD"] != DBNull.Value ? Convert.ToDecimal(x["CALIDAD"]) : 0,
                    Horario_Sesion1 = x["HORARIO_DIFERENCIA1"] != DBNull.Value ? Convert.ToDecimal(x["HORARIO_DIFERENCIA1"]) : 0,
                    Horario_Sesion2 = x["HORARIO_DIFERENCIA2"] != DBNull.Value ? Convert.ToDecimal(x["HORARIO_DIFERENCIA2"]) : 0,
                    Horario_Sesion3 = x["HORARIO_DIFERENCIA3"] != DBNull.Value ? Convert.ToDecimal(x["HORARIO_DIFERENCIA3"]) : 0,
                    Capacidad_Ordeño = x["CAPACIDADORDEÑO"] != DBNull.Value ? Convert.ToDecimal(x["CAPACIDADORDEÑO"]) : 0,
                    Porcentaje_Revueltas = x["PORCENREVUELTAS"] != DBNull.Value ? Convert.ToDecimal(x["PORCENREVUELTAS"]) : 0,
                    Litros_Sesion1 = x["LITROS1"] != DBNull.Value ? Convert.ToDecimal(x["LITROS1"]) : 0,
                    Litros_Sesion2 = x["LITROS2"] != DBNull.Value ? Convert.ToDecimal(x["LITROS2"]) : 0,
                    Litros_Sesion3 = x["LITROS3"] != DBNull.Value ? Convert.ToDecimal(x["LITROS3"]) : 0,
                    Horario_Dif_Sesion1 = x["HORARIO_DIF_FININI1"] != DBNull.Value ? Convert.ToDecimal(x["HORARIO_DIF_FININI1"]) : 0,
                    Horario_Dif_Sesion2 = x["HORARIO_DIF_FININI2"] != DBNull.Value ? Convert.ToDecimal(x["HORARIO_DIF_FININI2"]) : 0,
                    Horario_Dif_Sesion3 = x["HORARIO_DIF_FININI3"] != DBNull.Value ? Convert.ToDecimal(x["HORARIO_DIF_FININI3"]) : 0
                }).ToList();
            }
            catch (DbException ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "GetCalostroOrdeño", $"Error in {nameof(GetCalostroOrdeño)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private List<Mproduc> GetMediaProduccionNoIdReal(DateTime fechaInicio, DateTime fechaFin)
        {
            List<Mproduc> response = new List<Mproduc>();
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT  CDATE(m.FECHA)                                                    AS FechaG
                                        ,(LECFEDERAL + LECPLANTA)                                          AS TOTAL
                                        ,VACASORDEÑA                                                       AS ORDEÑO
                                        ,VACASSECAS                                                        AS SECAS
                                        ,VACASHATO                                                         AS HATO
                                        ,LACTOSA                                                           AS PLACT
                                        ,PROTEINA                                                          AS PPROT
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
                                FROM MPRODUC M
                                WHERE m.FECHA BETWEEN @julianaI AND @julianaF
                                ORDER BY m.FECHA ";

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@julianaI", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@julianaF", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new Mproduc()
                {
                    Fecha = x["FechaG"] != DBNull.Value ? Convert.ToDateTime(x["FechaG"]) : new DateTime(),
                    Total = x["TOTAL"] != DBNull.Value ? Convert.ToDecimal(x["TOTAL"]) : 0,
                    Vacas_Ordeño = x["ORDEÑO"] != DBNull.Value ? Convert.ToDecimal(x["ORDEÑO"]) : 0,
                    Vacas_Secas = x["SECAS"] != DBNull.Value ? Convert.ToDecimal(x["SECAS"]) : 0,
                    Vacas_Hato = x["HATO"] != DBNull.Value ? Convert.ToDecimal(x["HATO"]) : 0,
                    Lactosa = x["PLACT"] != DBNull.Value ? Convert.ToDecimal(x["PLACT"]) : 0,
                    Proteina = x["PPROT"] != DBNull.Value ? Convert.ToDecimal(x["PPROT"]) : 0,
                    Leche_Producida = x["LECPROD"] != DBNull.Value ? Convert.ToDecimal(x["LECPROD"]) : 0,
                    Antib = x["ANTIB"] != DBNull.Value ? Convert.ToDecimal(x["ANTIB"]) : 0,
                    Media = x["X"] != DBNull.Value ? Convert.ToDecimal(x["X"]) : 0,
                    Total_Producido = x["TOTALP"] != DBNull.Value ? Convert.ToDecimal(x["TOTALP"]) : 0,
                    Vaca_Antib = x["VACAANTIB"] != DBNull.Value ? Convert.ToDecimal(x["VACAANTIB"]) : 0,
                    ILCA_Venta = x["ILCA"] != DBNull.Value ? Convert.ToDecimal(x["ILCA"]) : 0,
                    IC_Venta = x["IC"] != DBNull.Value ? Convert.ToDecimal(x["IC"]) : 0,
                    EA = x["EA"] != DBNull.Value ? Convert.ToDecimal(x["EA"]) : 0,
                    ILCA_Produccion = x["ILCA_P"] != DBNull.Value ? Convert.ToDecimal(x["ILCA_P"]) : 0,
                    IC_Produccion = x["IC_P"] != DBNull.Value ? Convert.ToDecimal(x["IC_P"]) : 0,
                    Costo = x["COSTO"] != DBNull.Value ? Convert.ToDecimal(x["COSTO"]) : 0,
                    MH = x["MH"] != DBNull.Value ? Convert.ToDecimal(x["MH"]) : 0,
                    MS = x["MS"] != DBNull.Value ? Convert.ToDecimal(x["MS"]) : 0,
                    SA = x["SA"] != DBNull.Value ? Convert.ToDecimal(x["SA"]) : 0,
                    MSS = x["MSS"] != DBNull.Value ? Convert.ToDecimal(x["MSS"]) : 0,
                    EAS = x["EAS"] != DBNull.Value ? Convert.ToDecimal(x["EAS"]) : 0,
                    Noid_Ses1 = x["NOIDSES1"] != DBNull.Value ? Convert.ToDecimal(x["NOIDSES1"]) : 0,
                    Noid_Ses2 = x["NOIDSES2"] != DBNull.Value ? Convert.ToDecimal(x["NOIDSES2"]) : 0,
                    Noid_Ses3 = x["NOIDSES3"] != DBNull.Value ? Convert.ToDecimal(x["NOIDSES3"]) : 0
                }).ToList();
            }
            catch (DbException ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "GetMediaProduccionNoIdReal", $"Error in {nameof(GetMediaProduccionNoIdReal)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private List<Mproduc> GetMediaProduccionNoId(DateTime fechaInicio, DateTime fechaFin)
        {
            List<Mproduc> response = new List<Mproduc>();
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT  CDATE(m.FECHA)                                                    AS FechaG
                                        ,(LECFEDERAL + LECPLANTA)                                          AS TOTAL
                                        ,VACASORDEÑA                                                       AS ORDEÑO
                                        ,VACASSECAS                                                        AS SECAS
                                        ,VACASHATO                                                         AS HATO
                                        ,LACTOSA                                                           AS PLACT
                                        ,PROTEINA                                                          AS PPROT
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
                                        ,M.NOIDSES1                                                         AS NOIDSES1
                                        ,M.NOIDSES2                                                         AS NOIDSES2
                                        ,M.NOIDSES3                                                         AS NOIDSES3
                                FROM MPRODUC M
                                WHERE m.FECHA BETWEEN @julianaI AND @julianaF
                                ORDER BY m.FECHA ";

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@julianaI", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@julianaF", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new Mproduc()
                {
                    Fecha = x["FechaG"] != DBNull.Value ? Convert.ToDateTime(x["FechaG"]) : new DateTime(),
                    Total = x["TOTAL"] != DBNull.Value ? Convert.ToDecimal(x["TOTAL"]) : 0,
                    Vacas_Ordeño = x["ORDEÑO"] != DBNull.Value ? Convert.ToDecimal(x["ORDEÑO"]) : 0,
                    Vacas_Secas = x["SECAS"] != DBNull.Value ? Convert.ToDecimal(x["SECAS"]) : 0,
                    Vacas_Hato = x["HATO"] != DBNull.Value ? Convert.ToDecimal(x["HATO"]) : 0,
                    Lactosa = x["PLACT"] != DBNull.Value ? Convert.ToDecimal(x["PLACT"]) : 0,
                    Proteina = x["PPROT"] != DBNull.Value ? Convert.ToDecimal(x["PPROT"]) : 0,
                    Leche_Producida = x["LECPROD"] != DBNull.Value ? Convert.ToDecimal(x["LECPROD"]) : 0,
                    Antib = x["ANTIB"] != DBNull.Value ? Convert.ToDecimal(x["ANTIB"]) : 0,
                    Media = x["X"] != DBNull.Value ? Convert.ToDecimal(x["X"]) : 0,
                    Total_Producido = x["TOTALP"] != DBNull.Value ? Convert.ToDecimal(x["TOTALP"]) : 0,
                    Vaca_Antib = x["VACAANTIB"] != DBNull.Value ? Convert.ToDecimal(x["VACAANTIB"]) : 0,
                    ILCA_Venta = x["ILCA"] != DBNull.Value ? Convert.ToDecimal(x["ILCA"]) : 0,
                    IC_Venta = x["IC"] != DBNull.Value ? Convert.ToDecimal(x["IC"]) : 0,
                    EA = x["EA"] != DBNull.Value ? Convert.ToDecimal(x["EA"]) : 0,
                    ILCA_Produccion = x["ILCA_P"] != DBNull.Value ? Convert.ToDecimal(x["ILCA_P"]) : 0,
                    IC_Produccion = x["IC_P"] != DBNull.Value ? Convert.ToDecimal(x["IC_P"]) : 0,
                    Costo = x["COSTO"] != DBNull.Value ? Convert.ToDecimal(x["COSTO"]) : 0,
                    MH = x["MH"] != DBNull.Value ? Convert.ToDecimal(x["MH"]) : 0,
                    MS = x["MS"] != DBNull.Value ? Convert.ToDecimal(x["MS"]) : 0,
                    SA = x["SA"] != DBNull.Value ? Convert.ToDecimal(x["SA"]) : 0,
                    MSS = x["MSS"] != DBNull.Value ? Convert.ToDecimal(x["MSS"]) : 0,
                    EAS = x["EAS"] != DBNull.Value ? Convert.ToDecimal(x["EAS"]) : 0,
                    Noid_Ses1 = x["NOIDSES1"] != DBNull.Value ? Convert.ToDecimal(x["NOIDSES1"]) : 0,
                    Noid_Ses2 = x["NOIDSES2"] != DBNull.Value ? Convert.ToDecimal(x["NOIDSES2"]) : 0,
                    Noid_Ses3 = x["NOIDSES3"] != DBNull.Value ? Convert.ToDecimal(x["NOIDSES3"]) : 0
                }).ToList();
            }
            catch (DbException ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "GetMediaProduccionNoId", $"Error in {nameof(GetMediaProduccionNoId)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private PromedioHoja1 GetPromedioHoja1NoIdReal(DateTime fechaInicio, DateTime fechaFin)
        {
            PromedioHoja1 response = new PromedioHoja1();
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT  IIF(SUM(IIF(T3.TOTAL <> 0,1,0)) > 0 ,CLng(SUM(T3.TOTAL)/SUM(IIF(T3.TOTAL <> 0,1,0))),0)
                                        ,IIF(SUM(IIF(T3.ORDEÑO <> 0,1,0)) > 0 ,SUM(T3.ORDEÑO)/SUM(IIF(T3.ORDEÑO <> 0,1,0)),0)                AS ORDEÑO
                                        ,IIF(SUM(IIF(T3.SECAS <> 0,1,0)) > 0 ,SUM(T3.SECAS)/SUM(IIF(T3.SECAS <> 0,1,0)),0)                   AS SECAS
                                        ,IIF(SUM(IIF(T3.HATO <> 0,1,0)) > 0 ,SUM(T3.HATO)/SUM(IIF(T3.HATO <> 0,1,0)),0)                      AS HATO
                                        ,IIF(SUM(IIF(T3.PLACT <> 0,1,0)) > 0 ,SUM(T3.PLACT)/SUM(IIF(T3.PLACT <> 0,1,0)),0)                   AS LACT
                                        ,IIF(SUM(IIF(T3.PPROT <> 0,1,0)) > 0 ,SUM(T3.PPROT)/SUM(IIF(T3.PPROT <> 0,1,0)),0)                   AS PROT
                                        ,IIF(SUM(IIF(T3.UREA > 0,1,0)) > 0 ,SUM(T3.UREA)/SUM(IIF(T3.UREA > 0,1,0)),0)                      AS UREA
                                        ,IIF(SUM(IIF(T3.GRASA <> 0,1,0)) > 0 ,SUM(T3.GRASA)/SUM(IIF(T3.GRASA <> 0,1,0)),0)                   AS GRASA
                                        ,IIF(SUM(IIF(T3.CCS <> 0,1,0)) > 0 ,SUM(T3.CCS)/SUM(IIF(T3.CCS <> 0,1,0)),0)                         AS CCS
                                        ,IIF(SUM(IIF(T3.CTD <> 0,1,0)) > 0 ,CInt(SUM(T3.CTD)/SUM(IIF(T3.CTD <> 0,1,0))),0)                   AS CTD
                                        ,IIF(SUM(IIF(T3.LECPROD <> 0,1,0)) > 0 ,CLng(SUM(T3.LECPROD)/SUM(IIF(T3.LECPROD <> 0,1,0))),0)       AS LECHE
                                        ,IIF(SUM(IIF(T3.ANTIB <> 0,1,0)) > 0 ,CInt(SUM(T3.ANTIB)/SUM(IIF(T3.ANTIB <> 0,1,0))),0)             AS ANTIB
                                        ,IIF(SUM(IIF(T3.TOTALP <> 0,1,0)) > 0 ,CLng(SUM(T3.TOTALP)/SUM(IIF(T3.TOTALP <> 0,1,0))),0)          AS TOTAL2
                                        ,IIF(SUM(IIF(T3.DELORD <> 0,1,0)) > 0 ,CInt(SUM(T3.DELORD)/SUM(IIF(T3.DELORD <> 0,1,0))),0)          AS DEL
                                        ,IIF(SUM(IIF(T3.VACAANTIB <> 0,1,0)) > 0 ,CInt(SUM(T3.VACAANTIB)/SUM(IIF(T3.VACAANTIB <> 0,1,0))),0) AS ANT
                                        ,IIF(SUM(IIF(T3.criba1 <> 0,1,0)) > 0 ,CInt(SUM(T3.criba1)/SUM(IIF(T3.criba1 <> 0,1,0))),0)          AS N1
                                        ,IIF(SUM(IIF(T3.criba2 <> 0,1,0)) > 0 ,CInt(SUM(T3.criba2)/SUM(IIF(T3.criba2 <> 0,1,0))),0)          AS N2
                                        ,IIF(SUM(IIF(T3.criba3 <> 0,1,0)) > 0 ,CInt(SUM(T3.criba3)/SUM(IIF(T3.criba3 <> 0,1,0))),0)          AS N3
                                        ,IIF(SUM(IIF(T3.criba4 <> 0,1,0)) > 0 ,CInt(SUM(T3.criba4)/SUM(IIF(T3.criba4 <> 0,1,0))),0)          AS N4
                                        ,IIF(SUM(T3.ContNOIDSES1) > 0,CINT(SUM(T3.NOIDSES1)/SUM(T3.ContNOIDSES1)) ,0)                        AS SES1
                                        ,IIF(SUM(T3.ContNOIDSES2) > 0,CINT(SUM(T3.NOIDSES2)/SUM(T3.ContNOIDSES2)) ,0)                        AS SES2
                                        ,IIF(SUM(T3.ContNOIDSES3) > 0,CINT(SUM(T3.NOIDSES3)/SUM(T3.ContNOIDSES3)) ,0)                        AS SES3
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

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@inicio", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fin", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new PromedioHoja1()
                {
                    Ordeño = x["ORDEÑO"] != DBNull.Value ? Convert.ToDecimal(x["ORDEÑO"]) : 0,
                    Secas = x["SECAS"] != DBNull.Value ? Convert.ToDecimal(x["SECAS"]) : 0,
                    Hato = x["HATO"] != DBNull.Value ? Convert.ToDecimal(x["HATO"]) : 0,
                    Porcentaje_Lact = x["LACT"] != DBNull.Value ? Convert.ToDecimal(x["LACT"]) : 0,
                    Porcentaje_Prot = x["PROT"] != DBNull.Value ? Convert.ToDecimal(x["PROT"]) : 0,
                    Urea = x["UREA"] != DBNull.Value ? Convert.ToDecimal(x["UREA"]) : 0,
                    Porcentaje_Grasa = x["GRASA"] != DBNull.Value ? Convert.ToDecimal(x["GRASA"]) : 0,
                    CCS = x["CCS"] != DBNull.Value ? Convert.ToDecimal(x["CCS"]) : 0,
                    CTD = x["CTD"] != DBNull.Value ? Convert.ToDecimal(x["CTD"]) : 0,
                    Leche = x["LECHE"] != DBNull.Value ? Convert.ToDecimal(x["LECHE"]) : 0,
                    Antib = x["ANTIB"] != DBNull.Value ? Convert.ToDecimal(x["ANTIB"]) : 0,
                    Total = x["TOTAL2"] != DBNull.Value ? Convert.ToDecimal(x["TOTAL2"]) : 0,
                    DEL = x["DEL"] != DBNull.Value ? Convert.ToDecimal(x["DEL"]) : 0,
                    Ant = x["ANT"] != DBNull.Value ? Convert.ToDecimal(x["ANT"]) : 0,
                    Cribas_N1 = x["N1"] != DBNull.Value ? Convert.ToDecimal(x["N1"]) : 0,
                    Cribas_N2 = x["N2"] != DBNull.Value ? Convert.ToDecimal(x["N2"]) : 0,
                    Cribas_N3 = x["N3"] != DBNull.Value ? Convert.ToDecimal(x["N3"]) : 0,
                    Cribas_N4 = x["N4"] != DBNull.Value ? Convert.ToDecimal(x["N4"]) : 0,
                    Sesion1 = x["SES1"] != DBNull.Value ? Convert.ToDecimal(x["SES1"]) : 0,
                    Sesion2 = x["SES2"] != DBNull.Value ? Convert.ToDecimal(x["SES2"]) : 0,
                    Sesion3 = x["SES3"] != DBNull.Value ? Convert.ToDecimal(x["SES3"]) : 0
                }).ToList().FirstOrDefault();
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "PromedioHoja1NoIdReal", $"Error in {nameof(GetPromedioHoja1NoIdReal)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }


            return response;
        }

        private PromedioHoja1 GetPromedioHoja1NoId(DateTime fechaInicio, DateTime fechaFin)
        {
            PromedioHoja1 response = new PromedioHoja1();
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT  IIF(SUM(IIF(T3.ORDEÑO <> 0,1,0)) > 0 ,SUM(T3.ORDEÑO)/SUM(IIF(T3.ORDEÑO <> 0,1,0)),0)                AS ORDEÑO
                                        ,IIF(SUM(IIF(T3.SECAS <> 0,1,0)) > 0 ,SUM(T3.SECAS)/SUM(IIF(T3.SECAS <> 0,1,0)),0)                   AS SECAS
                                        ,IIF(SUM(IIF(T3.HATO <> 0,1,0)) > 0 ,SUM(T3.HATO)/SUM(IIF(T3.HATO <> 0,1,0)),0)                      AS HATO
                                        ,IIF(SUM(IIF(T3.PLACT <> 0,1,0)) > 0 ,SUM(T3.PLACT)/SUM(IIF(T3.PLACT <> 0,1,0)),0)                   AS LACT
                                        ,IIF(SUM(IIF(T3.PPROT <> 0,1,0)) > 0 ,SUM(T3.PPROT)/SUM(IIF(T3.PPROT <> 0,1,0)),0)                   AS PROT
                                        ,IIF(SUM(IIF(T3.UREA > 0,1,0)) > 0 ,SUM(T3.UREA)/SUM(IIF(T3.UREA > 0,1,0)),0)                      AS UREA
                                        ,IIF(SUM(IIF(T3.GRASA <> 0,1,0)) > 0 ,SUM(T3.GRASA)/SUM(IIF(T3.GRASA <> 0,1,0)),0)                   AS GRASA
                                        ,IIF(SUM(IIF(T3.CCS <> 0,1,0)) > 0 ,SUM(T3.CCS)/SUM(IIF(T3.CCS <> 0,1,0)),0)                         AS CCS
                                        ,IIF(SUM(IIF(T3.CTD <> 0,1,0)) > 0 ,CInt(SUM(T3.CTD)/SUM(IIF(T3.CTD <> 0,1,0))),0)                   AS CTD
                                        ,IIF(SUM(IIF(T3.LECPROD <> 0,1,0)) > 0 ,CLng(SUM(T3.LECPROD)/SUM(IIF(T3.LECPROD <> 0,1,0))),0)       AS LECHE
                                        ,IIF(SUM(IIF(T3.ANTIB <> 0,1,0)) > 0 ,CInt(SUM(T3.ANTIB)/SUM(IIF(T3.ANTIB <> 0,1,0))),0)             AS ANTIB
                                        ,IIF(SUM(IIF(T3.TOTALP <> 0,1,0)) > 0 ,CLng(SUM(T3.TOTALP)/SUM(IIF(T3.TOTALP <> 0,1,0))),0)          AS TOTAL2
                                        ,IIF(SUM(IIF(T3.DELORD <> 0,1,0)) > 0 ,CInt(SUM(T3.DELORD)/SUM(IIF(T3.DELORD <> 0,1,0))),0)          AS DEL
                                        ,IIF(SUM(IIF(T3.VACAANTIB <> 0,1,0)) > 0 ,CInt(SUM(T3.VACAANTIB)/SUM(IIF(T3.VACAANTIB <> 0,1,0))),0) AS ANT
                                        ,IIF(SUM(IIF(T3.criba1 <> 0,1,0)) > 0 ,CInt(SUM(T3.criba1)/SUM(IIF(T3.criba1 <> 0,1,0))),0)          AS N1
                                        ,IIF(SUM(IIF(T3.criba2 <> 0,1,0)) > 0 ,CInt(SUM(T3.criba2)/SUM(IIF(T3.criba2 <> 0,1,0))),0)          AS N2
                                        ,IIF(SUM(IIF(T3.criba3 <> 0,1,0)) > 0 ,CInt(SUM(T3.criba3)/SUM(IIF(T3.criba3 <> 0,1,0))),0)          AS N3
                                        ,IIF(SUM(IIF(T3.criba4 <> 0,1,0)) > 0 ,CInt(SUM(T3.criba4)/SUM(IIF(T3.criba4 <> 0,1,0))),0)          AS N4
                                        ,IIF(SUM(T3.ContNOIDSES1) > 0,CINT(SUM(T3.NOIDSES1)/SUM(T3.ContNOIDSES1)) ,0)                        AS SES1
                                        ,IIF(SUM(T3.ContNOIDSES2) > 0,CINT(SUM(T3.NOIDSES2)/SUM(T3.ContNOIDSES2)) ,0)                        AS SES2
                                        ,IIF(SUM(T3.ContNOIDSES3) > 0,CINT(SUM(T3.NOIDSES3)/SUM(T3.ContNOIDSES3)) ,0)                        AS SES3
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
				                                        ,M.NOIDSES1
				                                        ,M.NOIDSES2
				                                        ,M.NOIDSES3
				                                        ,IIF(M.NOIDSES1 > 0,1,0)                                           AS ContNOIDSES1
				                                        ,IIF(M.NOIDSES2 > 0,1,0)                                           AS ContNOIDSES2
				                                        ,IIF(M.NOIDSES3 > 0,1,0)                                           AS ContNOIDSES3
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

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@inicio", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fin", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new PromedioHoja1()
                {
                    Ordeño = x["ORDEÑO"] != DBNull.Value ? Convert.ToDecimal(x["ORDEÑO"]) : 0,
                    Secas = x["SECAS"] != DBNull.Value ? Convert.ToDecimal(x["SECAS"]) : 0,
                    Hato = x["HATO"] != DBNull.Value ? Convert.ToDecimal(x["HATO"]) : 0,
                    Porcentaje_Lact = x["LACT"] != DBNull.Value ? Convert.ToDecimal(x["LACT"]) : 0,
                    Porcentaje_Prot = x["PROT"] != DBNull.Value ? Convert.ToDecimal(x["PROT"]) : 0,
                    Urea = x["UREA"] != DBNull.Value ? Convert.ToDecimal(x["UREA"]) : 0,
                    Porcentaje_Grasa = x["GRASA"] != DBNull.Value ? Convert.ToDecimal(x["GRASA"]) : 0,
                    CCS = x["CCS"] != DBNull.Value ? Convert.ToDecimal(x["CCS"]) : 0,
                    CTD = x["CTD"] != DBNull.Value ? Convert.ToDecimal(x["CTD"]) : 0,
                    Leche = x["LECHE"] != DBNull.Value ? Convert.ToDecimal(x["LECHE"]) : 0,
                    Antib = x["ANTIB"] != DBNull.Value ? Convert.ToDecimal(x["ANTIB"]) : 0,
                    Total = x["TOTAL2"] != DBNull.Value ? Convert.ToDecimal(x["TOTAL2"]) : 0,
                    DEL = x["DEL"] != DBNull.Value ? Convert.ToDecimal(x["DEL"]) : 0,
                    Ant = x["ANT"] != DBNull.Value ? Convert.ToDecimal(x["ANT"]) : 0,
                    Cribas_N1 = x["N1"] != DBNull.Value ? Convert.ToDecimal(x["N1"]) : 0,
                    Cribas_N2 = x["N2"] != DBNull.Value ? Convert.ToDecimal(x["N2"]) : 0,
                    Cribas_N3 = x["N3"] != DBNull.Value ? Convert.ToDecimal(x["N3"]) : 0,
                    Cribas_N4 = x["N4"] != DBNull.Value ? Convert.ToDecimal(x["N4"]) : 0,
                    Sesion1 = x["SES1"] != DBNull.Value ? Convert.ToDecimal(x["SES1"]) : 0,
                    Sesion2 = x["SES2"] != DBNull.Value ? Convert.ToDecimal(x["SES2"]) : 0,
                    Sesion3 = x["SES3"] != DBNull.Value ? Convert.ToDecimal(x["SES3"]) : 0
                }).ToList().FirstOrDefault();
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "PromedioHoja1NoId", $"Error in {nameof(GetPromedioHoja1NoId)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }


            return response;
        }

        private CalostroYOrdeña GetPromedioCalostro(DateTime fechaInicio, DateTime fechaFin)
        {
            CalostroYOrdeña response = new CalostroYOrdeña() { Fecha = fechaFin };
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT  IIF(SUM(ESPERADO_VQ + ESPERADO_VC) > 0,SUM(DIFERENCIA)/SUM(ESPERADO_VQ + ESPERADO_VC)*100,0)                AS Porcentaje
									   ,AVG(DIFERENCIA)                                                                                             AS DIFERENCIA_
									   ,IIF(SUM(IIF((CALIDAD) <> 0,1,0)) > 0,SUM(CALIDAD) /SUM(IIF((CALIDAD) <> 0,1,0)),0)                          AS CALIDAD_
									   ,AVG(HORARIO_DIFERENCIA1)                                                                                    AS HORARIO_DIFERENCIA1_
									   ,AVG(HORARIO_DIFERENCIA2)                                                                                    AS HORARIO_DIFERENCIA2_
									   ,AVG(HORARIO_DIFERENCIA3)                                                                                    AS HORARIO_DIFERENCIA3_
									   ,IIF(SUM(IIF((CAPACIDADORDEÑO) <> 0,1,0)) > 0,SUM(CAPACIDADORDEÑO) /SUM(IIF((CAPACIDADORDEÑO) <> 0,1,0)) ,0) AS CAPACIDADORDEÑO_
									   ,IIF(SUM(IIF(PORCENREVUELTAS <> 0,1,0)) > 0,SUM(PORCENREVUELTAS)/SUM(IIF(PORCENREVUELTAS <> 0,1,0)),0)       AS PORCENREVUELTAS_
									   ,IIF(SUM(IIF((LITROS1) <> 0,1,0)) > 0,SUM(LITROS1)/SUM(IIF((LITROS1) <> 0,1,0)),0)                           AS LITROS1_
									   ,IIF(SUM(IIF((LITROS2) <> 0,1,0)) > 0,SUM(LITROS2)/SUM(IIF((LITROS2) <> 0,1,0)),0)                           AS LITROS2_
									   ,IIF(SUM(IIF((LITROS3) <> 0,1,0)) > 0,SUM(LITROS3)/SUM(IIF((LITROS3) <> 0,1,0)),0)                           AS LITROS3_
								FROM CALOSTROYORDEÑODIA
						WHERE FECHA BETWEEN @fechaInicio AND @fechaFin";

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@fechaInicio", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new CalostroYOrdeña()
                {
                    Fecha = fechaFin,
                    Porcentaje = x["Porcentaje"] != DBNull.Value ? Convert.ToDecimal(x["Porcentaje"]) : 0,
                    Diferencia = x["DIFERENCIA_"] != DBNull.Value ? Convert.ToDecimal(x["DIFERENCIA_"]) : 0,
                    Calidad = x["CALIDAD_"] != DBNull.Value ? Convert.ToDecimal(x["CALIDAD_"]) : 0,
                    Horario_Sesion1 = x["HORARIO_DIFERENCIA1_"] != DBNull.Value ? Math.Round(Convert.ToDecimal(x["HORARIO_DIFERENCIA1_"])) : 0,
                    Horario_Sesion2 = x["HORARIO_DIFERENCIA2_"] != DBNull.Value ? Math.Round(Convert.ToDecimal(x["HORARIO_DIFERENCIA2_"])) : 0,
                    Horario_Sesion3 = x["HORARIO_DIFERENCIA3_"] != DBNull.Value ? Math.Round(Convert.ToDecimal(x["HORARIO_DIFERENCIA3_"])) : 0,
                    Capacidad_Ordeño = x["CAPACIDADORDEÑO_"] != DBNull.Value ? Convert.ToDecimal(x["CAPACIDADORDEÑO_"]) : 0,
                    Porcentaje_Revueltas = x["PORCENREVUELTAS_"] != DBNull.Value ? Math.Round(Convert.ToDecimal(x["PORCENREVUELTAS_"]), 1) : 0,
                    Litros_Sesion1 = x["LITROS1_"] != DBNull.Value ? Convert.ToDecimal(x["LITROS1_"]) : 0,
                    Litros_Sesion2 = x["LITROS2_"] != DBNull.Value ? Convert.ToDecimal(x["LITROS2_"]) : 0,
                    Litros_Sesion3 = x["LITROS3_"] != DBNull.Value ? Convert.ToDecimal(x["LITROS3_"]) : 0
                }).ToList().FirstOrDefault();
            }
            catch (DbException ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "PromedioCalostro", $"Error in {nameof(GetPromedioCalostro)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private DatosProduccion GetTotalDatosDeProduccion(DateTime fechaInicio, DateTime fechaFin)
        {
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);
            DatosProduccion response = new DatosProduccion();

            try
            {
                string query = @"SELECT  IIF(SUM(ContUrea1_) > 0,SUM(Urea1_)/SUM(ContUrea1_),0)   AS Urea1
                                       ,IIF(SUM(ContGrasa1_) > 0,SUM(Grasa1_)/SUM(ContGrasa1_),0) AS Grasa1
                                       ,IIF(SUM(ContCCS1_) > 0,SUM(CCS1_)/SUM(ContCCS1_),0)       AS CCS1
                                       ,IIF(SUM(ContCTD1_) > 0,SUM(CTD1_)/SUM(ContCTD1_),0)       AS CTD1
                            FROM
                            (
	                            SELECT  CDATE(RESULTADOS.FECHA)        AS FechaG
	                                    ,AVG(RESULTADOS.PROTES)         AS Proteina1_
	                                    ,SUM(RESULTADOS.UREAS)          AS Urea1_
	                                    ,IIF(SUM(RESULTADOS.UREAS) > 0,1,0)  AS ContUrea1_
	                                    ,SUM(RESULTADOS.GRASAS)         AS Grasa1_
	                                    ,IIF(SUM(RESULTADOS.GRASAS) > 0,1,0) AS ContGrasa1_
	                                    ,SUM(RESULTADOS.CCSS)           AS CCS1_
	                                    ,IIF(SUM(RESULTADOS.CCSS)  > 0,1,0)   AS ContCCS1_
	                                    ,SUM(RESULTADOS.CTDS)           AS CTD1_
	                                    ,IIF(SUM(RESULTADOS.CTDS) > 0,1,0)   AS ContCTD1_
	                            FROM
	                            (
		                            SELECT  LECHEXDIA.FECHA
		                                   ,VALORES.PROTEINA                                                                      AS PROTES
		                                   ,IIF(ISNULL(LECHEXDIA.LG),NULL,(VALORES.LITROSXTANQUE / LECHEXDIA.LG) * VALORES.GRASA) AS GRASAS
		                                   ,IIF(ISNULL(LECHEXDIA.LU),NULL,(VALORES.LITROSXTANQUE / LECHEXDIA.LU) * VALORES.UREA)  AS UREAS
		                                   ,IIF(ISNULL(LECHEXDIA.LCC),NULL,(VALORES.LITROSXTANQUE / LECHEXDIA.LCC) * VALORES.CCS) AS CCSS
		                                   ,IIF(ISNULL(LECHEXDIA.LCT),NULL,(VALORES.LITROSXTANQUE / LECHEXDIA.LCT) * VALORES.CTD) AS CTDS
		                            FROM
		                            (
			                            SELECT  LITROSxDIA.DIA             AS FECHA
			                                   ,SUM(LITROSxDIA.LITROGRASA) AS LG
			                                   ,SUM(LITROSxDIA.LITROUREA)  AS LU
			                                   ,SUM(LITROSxDIA.LITROCCS)   AS LCC
			                                   ,SUM(LITROSxDIA.LITROCTS)   AS LCT
			                            FROM
			                            (
				                            SELECT  FECHA                        AS DIA
				                                   ,GRASA
				                                   ,IIF(GRASA > 0,LITROSXTANQUE) AS LITROGRASA
				                                   ,IIF(UREA > 0,LITROSXTANQUE)  AS LITROUREA
				                                   ,IIF(CCS > 0,LITROSXTANQUE)   AS LITROCCS
				                                   ,IIF(CTD > 0,LITROSXTANQUE)   AS LITROCTS
				                            FROM dproduc
				                            WHERE FECHA BETWEEN @julianaI AND @julianaF
				                            ORDER BY FECHA 
			                            ) LITROSxDIA
			                            GROUP BY  LITROSxDIA.DIA
		                            ) LECHEXDIA
		                            LEFT JOIN
		                            (
			                            SELECT  FECHA
			                                   ,PROTEINA
			                                   ,LITROSXTANQUE
			                                   ,GRASA
			                                   ,UREA
			                                   ,CCS
			                                   ,CTD
			                            FROM dproduc
			                            WHERE FECHA BETWEEN @julianaI AND @julianaF
			                            ORDER BY FECHA 
		                            ) VALORES
		                            ON VALORES.FECHA = LECHEXDIA.FECHA
	                            )RESULTADOS
	                            GROUP BY  RESULTADOS.FECHA
                            )Datos";

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@julianaI", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@julianaF", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new DatosProduccion()
                {
                    Proteina = 0,
                    Urea = x["Urea1"] != DBNull.Value ? Convert.ToDecimal(x["Urea1"]) : 0,
                    Grasa = x["Grasa1"] != DBNull.Value ? Convert.ToDecimal(x["Grasa1"]) : 0,
                    CCS = x["CCS1"] != DBNull.Value ? Convert.ToDecimal(x["CCS1"]) : 0,
                    CTD = x["CTD1"] != DBNull.Value ? Convert.ToDecimal(x["CTD1"]) : 0
                }).ToList().FirstOrDefault();

            }
            catch (DbException ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "TotalDatosDeProduccion", $"Error in {nameof(GetTotalDatosDeProduccion)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private bool EsPrecioLecheFacturado(DateTime fechaPeriodo)
        {
            DateTime fechaActual = DateTime.Today;
            DateTime fechaMesAnt = DateTime.Today.AddMonths(-1);


            if (fechaActual.Month == fechaPeriodo.Month && fechaActual.Year == fechaPeriodo.Year)
                return false;
            else if (fechaMesAnt.Year == fechaPeriodo.Year && fechaMesAnt.Month == fechaPeriodo.Month && fechaActual.Day < 7)
                return false;
            else
                return true;

        }

        private decimal GetPrecioLecheFacturado(DateTime fechaInicio, DateTime fechaFin)
        {
            decimal precio = 0;
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT AVG(PRECIOLECHE) AS PRECIO_LECHE
                                 FROM MPRODUC m
                                 WHERE m.FECHA BETWEEN @fechaInicio AND @fechaFin";

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@fechaInicio", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                DataTable dt = db.EjecutarConsultaTabla();

                if (dt.Rows.Count > 0)
                    precio = dt.Rows[0]["PRECIO_LECHE"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["PRECIO_LECHE"]) : 0;
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "PrecioLecheFacturado", $"Error in {nameof(GetPrecioLecheFacturado)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return precio;
        }

        private decimal GetPrecioLecheCalculado(int centroId, DateTime fechaInicio, DateTime fechaFin)
        {
            decimal precio = 0;
            ModeloDatosSQL db = new ModeloDatosSQL(_sqlConnection);

            try
            {
                DateTime hoy = DateTime.Today;
                DateTime mesAnt = hoy.AddMonths(-1);
                string query = "";

                db.Conectar();
                if (fechaInicio.Year == mesAnt.Year && fechaInicio.Month == mesAnt.Month && fechaFin.Month == mesAnt.Month && fechaFin.Year == mesAnt.Year && hoy.Day < 7)
                {
                    query = @"WITH FechaPrecioLeche AS
                            (
                                SELECT Centro, MAX(Fecha) AS Fecha
                                FROM DBCONSUMO.dbo.PrecioLecheTbl
                                WHERE Centro = @rancho
                                AND CONVERT(int, FORMAT(Fecha, 'yyyyMM')) = @fecha
                                GROUP BY Centro
                            )
                            select p.PrecioLeche AS Precio_Leche
                            FROM DBCONSUMO.dbo.PrecioLecheTbl p 
                            INNER JOIN FechaPrecioLeche fpl ON p.Centro = fpl.Centro AND p.Fecha = fpl.Fecha";

                    db.CrearComando(query, tipoComandoSQL.query);
                    db.AsignarParametro("@rancho", centroId);
                    db.AsignarParametro("@fecha", Convert.ToInt32(fechaFin.ToString("yyyyMM")));
                }
                else
                {
                    query = @"WITH FechaPrecioLeche AS
                            (
                                SELECT Centro, MAX(Fecha) AS Fecha
                                FROM DBCONSUMO.dbo.PrecioLecheTbl
                                WHERE Centro = @rancho
                                GROUP BY Centro
                            )
                            select p.PrecioLeche AS Precio_Leche
                            FROM DBCONSUMO.dbo.PrecioLecheTbl p 
                            INNER JOIN FechaPrecioLeche fpl ON p.Centro = fpl.Centro AND p.Fecha = fpl.Fecha";

                    db.CrearComando(query, tipoComandoSQL.query);
                    db.AsignarParametro("@rancho", centroId);
                }

                DataTable dt = db.EjecutarConsultaTabla();

                if (dt.Rows.Count > 0)
                    precio = dt.Rows[0]["Precio_Leche"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["Precio_Leche"]) : 0;

            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "PrecioLecheCalculado", $"Error in {nameof(GetPrecioLecheCalculado)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return precio;
        }

        private List<InventarioAfiXDia> GetInventariosAFI(DateTime fechaInicio, DateTime fechaFin)
        {
            List<InventarioAfiXDia> response = new List<InventarioAfiXDia>();
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT  CDATE(fecha)                                                                                 AS FechaG
                                        ,jaulas                                                                                       AS JAULAS_
                                        ,destetadas                                                                                   AS CRECIMIENTO
                                        ,destetadas2                                                                                  AS DESARROLLO
                                        ,vaquillas                                                                                    AS VAQUILLAS_
                                        ,vacassecas                                                                                   AS SECAS
                                        ,vacasordeña                                                                                  AS PRODUCCION
                                        ,(vqreto + vcreto)                                                                            AS RETO
                                        ,(jaulas + destetadas + destetadas2 + vaquillas)                                              AS Crianza
                                        ,(jaulas + destetadas + destetadas2 + vaquillas + vacassecas + vacasordeña + vqreto + vcreto) AS InventarioTotal
                                FROM INVENTARIOAFIXDIA
                                WHERE fecha BETWEEN @fechaInicio AND @fechaFin
                                ORDER BY 1";

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@fechaInicio", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new InventarioAfiXDia()
                {
                    Fecha = x["FechaG"] != DBNull.Value ? Convert.ToDateTime(x["FechaG"]) : new DateTime(),
                    Ordeño = x["PRODUCCION"] != DBNull.Value ? Convert.ToDecimal(x["PRODUCCION"]) : 0,
                    Secas = x["SECAS"] != DBNull.Value ? Convert.ToDecimal(x["SECAS"]) : 0,
                    Reto = x["RETO"] != DBNull.Value ? Convert.ToDecimal(x["RETO"]) : 0,
                    Jaulas = x["JAULAS_"] != DBNull.Value ? Convert.ToDecimal(x["JAULAS_"]) : 0,
                    Crecimiento = x["CRECIMIENTO"] != DBNull.Value ? Convert.ToDecimal(x["CRECIMIENTO"]) : 0,
                    Desarrollo = x["DESARROLLO"] != DBNull.Value ? Convert.ToDecimal(x["DESARROLLO"]) : 0,
                    Vaquillas = x["VAQUILLAS_"] != DBNull.Value ? Convert.ToDecimal(x["VAQUILLAS_"]) : 0,
                    Crianza = x["Crianza"] != DBNull.Value ? Convert.ToDecimal(x["Crianza"]) : 0,
                    InventarioTotal = x["InventarioTotal"] != DBNull.Value ? Convert.ToDecimal(x["InventarioTotal"]) : 0
                }).ToList();
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "GetInventariosAFI", $"Error in {nameof(GetInventariosAFI)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }


            return response;
        }

        private List<Mproduc2> GetMediaProduccion2(DateTime fechaInicio, DateTime fechaFin)
        {
            List<Mproduc2> response = new List<Mproduc2>();
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT  CDATE(FECHA)                                                    AS FECHAG
                                        ,m.COSTO                                                         AS Ordeño_Costo
                                        ,m.JAULAS                                                        AS Jaulas_Costo
                                        ,m.MH_DI                                                         AS Crecimiento_MH
                                        ,m.BEC1                                                          AS Crecimiento_Costo
                                        ,m.MS_DI * 100                                                   AS Crecimiento_PorcentajeMS
                                        ,m.MH_DI * m.MS_DI                                               AS Crecimiento_MS
                                        ,IIf((m.MH_DI * m.MS_DI) > 0,M.BEC1 / (m.MH_DI * m.MS_DI),0)     AS Crecimiento_CostoMS
                                        ,m.MH_DII                                                        AS Desarrollo_MH
                                        ,m.BEC2                                                          AS Desarrollo_Costo
                                        ,m.MS_DII * 100                                                  AS Desarrollo_PorcentajeMS
                                        ,m.MH_DII * m.MS_DII                                             AS Desarrollo_MS
                                        ,IIf((m.MH_DII * m.MS_DII) > 0,M.BEC2 / (m.MH_DII * m.MS_DII),0) AS Desarrollo_CostoMS
                                        ,m.MH_VQ                                                         AS Vaquillas_MH
                                        ,m.VAQ                                                           AS Vaquillas_Costo
                                        ,m.MS_VQ * 100                                                   AS Vaquillas_PorcentajeMS
                                        ,m.MH_VQ * m.MS_VQ                                               AS Vaquillas_MS
                                        ,IIf((m.MH_VQ * m.MS_VQ) > 0,m.VAQ / (m.MH_VQ * m.MS_VQ),0)      AS Vaquillas_CostoMS
                                        ,m.MH_S                                                          AS Secas_MH
                                        ,m.MS_S * 100                                                    AS Secas_PorcentajeMS
                                        ,m.MH_S * m.MS_S                                                 AS Secas_MS
                                        ,m.SA_S                                                          AS Secas_SA
                                        ,(m.MH_S - m.SA_S) * m.MS_S                                      AS Secas_MSS
                                        ,IIf(m.MH_S > 0,m.SA_S / m.MH_S * 100,0)                         AS Secas_PorcentajeSob
                                        ,m.SEC                                                           AS Secas_Costo
                                        ,IIf(m.MS_S > 0 AND m.MH_S > 0,m.SEC / (m.MS_S * m.MH_S),0)      AS Secas_CostoMS
                                        ,m.MH_R                                                          AS Reto_MH
                                        ,m.MS_R * 100                                                    AS Reto_PorcentajeMS
                                        ,m.MH_R * m.MS_R                                                 AS Reto_MS
                                        ,m.SA_R                                                          AS Reto_SA
                                        ,(m.MH_R - m.SA_R) * m.MS_R                                      AS Reto_MSS
                                        ,IIF(m.MH_R > 0,m.SA_R / m.MH_R * 100,0)                         AS Reto_PorcentajeSob
                                        ,m.RETO                                                          AS Reto_Costo
                                        ,IIF(m.MS_R > 0 AND m.MH_R > 0,m.RETO / (m.MS_R * m.MH_R),0)     AS Reto_CostoMS
                                        ,m.LECPROD                                                       AS LecheProd
                                FROM MPRODUC m
                                WHERE m.FECHA BETWEEN @fechaInicio AND @fechaFin ";

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@fechaInicio", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new Mproduc2()
                {
                    Fecha = x["FECHAG"] != DBNull.Value ? Convert.ToDateTime(x["FECHAG"]) : new DateTime(),
                    Ordeño_Costo = x["Ordeño_Costo"] != DBNull.Value ? Convert.ToDecimal(x["Ordeño_Costo"]) : 0,
                    Jaulas_Costo = x["Jaulas_Costo"] != DBNull.Value ? Convert.ToDecimal(x["Jaulas_Costo"]) : 0,
                    Crecimiento_MH = x["Crecimiento_MH"] != DBNull.Value ? Convert.ToDecimal(x["Crecimiento_MH"]) : 0,
                    Crecimiento_Porcentaje_MS = x["Crecimiento_PorcentajeMS"] != DBNull.Value ? Convert.ToDecimal(x["Crecimiento_PorcentajeMS"]) : 0,
                    Crecimiento_MS = x["Crecimiento_MS"] != DBNull.Value ? Convert.ToDecimal(x["Crecimiento_MS"]) : 0,
                    Crecimiento_Costo = x["Crecimiento_Costo"] != DBNull.Value ? Convert.ToDecimal(x["Crecimiento_Costo"]) : 0,
                    Crecimiento_Costo_MS = x["Crecimiento_CostoMS"] != DBNull.Value ? Convert.ToDecimal(x["Crecimiento_CostoMS"]) : 0,
                    Desarrollo_MH = x["Desarrollo_MH"] != DBNull.Value ? Convert.ToDecimal(x["Desarrollo_MH"]) : 0,
                    Desarrollo_Porcentaje_MS = x["Desarrollo_PorcentajeMS"] != DBNull.Value ? Convert.ToDecimal(x["Desarrollo_PorcentajeMS"]) : 0,
                    Desarrollo_MS = x["Desarrollo_MS"] != DBNull.Value ? Convert.ToDecimal(x["Desarrollo_MS"]) : 0,
                    Desarrollo_Costo = x["Desarrollo_Costo"] != DBNull.Value ? Convert.ToDecimal(x["Desarrollo_Costo"]) : 0,
                    Desarrollo_Costo_MS = x["Desarrollo_CostoMS"] != DBNull.Value ? Convert.ToDecimal(x["Desarrollo_CostoMS"]) : 0,
                    Vaquillas_MH = x["Vaquillas_MH"] != DBNull.Value ? Convert.ToDecimal(x["Vaquillas_MH"]) : 0,
                    Vaquillas_Porcentaje_MS = x["Vaquillas_PorcentajeMS"] != DBNull.Value ? Convert.ToDecimal(x["Vaquillas_PorcentajeMS"]) : 0,
                    Vaquillas_MS = x["Vaquillas_MS"] != DBNull.Value ? Convert.ToDecimal(x["Vaquillas_MS"]) : 0,
                    Vaquillas_Costo = x["Vaquillas_Costo"] != DBNull.Value ? Convert.ToDecimal(x["Vaquillas_Costo"]) : 0,
                    Vaquillas_Costo_MS = x["Vaquillas_CostoMS"] != DBNull.Value ? Convert.ToDecimal(x["Vaquillas_CostoMS"]) : 0,
                    Reto_MH = x["Reto_MH"] != DBNull.Value ? Convert.ToDecimal(x["Reto_MH"]) : 0,
                    Reto_Porcentaje_MS = x["Reto_PorcentajeMS"] != DBNull.Value ? Convert.ToDecimal(x["Reto_PorcentajeMS"]) : 0,
                    Reto_MS = x["Reto_MS"] != DBNull.Value ? Convert.ToDecimal(x["Reto_MS"]) : 0,
                    Reto_SA = x["Reto_SA"] != DBNull.Value ? Convert.ToDecimal(x["Reto_SA"]) : 0,
                    Reto_MSS = x["Reto_MSS"] != DBNull.Value ? Convert.ToDecimal(x["Reto_MSS"]) : 0,
                    Reto_Porcentaje_SA = x["Reto_PorcentajeSob"] != DBNull.Value ? Convert.ToDecimal(x["Reto_PorcentajeSob"]) : 0,
                    Reto_Costo = x["Reto_Costo"] != DBNull.Value ? Convert.ToDecimal(x["Reto_Costo"]) : 0,
                    Reto_Costo_MS = x["Reto_CostoMS"] != DBNull.Value ? Convert.ToDecimal(x["Reto_CostoMS"]) : 0,
                    Secas_MH = x["Secas_MH"] != DBNull.Value ? Convert.ToDecimal(x["Secas_MH"]) : 0,
                    Secas_Porcentaje_MS = x["Secas_PorcentajeMS"] != DBNull.Value ? Convert.ToDecimal(x["Secas_PorcentajeMS"]) : 0,
                    Secas_MS = x["Secas_MS"] != DBNull.Value ? Convert.ToDecimal(x["Secas_MS"]) : 0,
                    Secas_SA = x["Secas_SA"] != DBNull.Value ? Convert.ToDecimal(x["Secas_SA"]) : 0,
                    Secas_MSS = x["Secas_MSS"] != DBNull.Value ? Convert.ToDecimal(x["Secas_MSS"]) : 0,
                    Secas_Porcentaje_SA = x["Secas_PorcentajeSob"] != DBNull.Value ? Convert.ToDecimal(x["Secas_PorcentajeSob"]) : 0,
                    Secas_Costo = x["Secas_Costo"] != DBNull.Value ? Convert.ToDecimal(x["Secas_Costo"]) : 0,
                    Secas_Costo_MS = x["Secas_CostoMS"] != DBNull.Value ? Convert.ToDecimal(x["Secas_CostoMS"]) : 0,
                    LecheProd = x["LecheProd"] != DBNull.Value ? Convert.ToDecimal(x["LecheProd"]) : 0
                }).ToList();
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "MediaProduccion2", $"Error in {nameof(GetMediaProduccion2)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private InventarioAfiXDia GetPromedioInventariosAFI(DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            InventarioAfiXDia response = new InventarioAfiXDia();
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);
            mensaje = string.Empty;

            try
            {
                string query = @"SELECT  Round(SUM(jaulas)/(@fechaFin - @fechaInicio + 1),0)                                                                                       AS JAULAS_
                                        ,Round(SUM(destetadas)/(@fechaFin - @fechaInicio + 1),0)                                                                                   AS CRECIMIENTO
                                        ,Round(SUM(destetadas2)/(@fechaFin - @fechaInicio + 1),0)                                                                                  AS DESARROLLO
                                        ,Round(SUM(vaquillas)/(@fechaFin - @fechaInicio + 1),0)                                                                                    AS VAQUILLAS_
                                        ,Round(SUM(vacassecas)/(@fechaFin - @fechaInicio + 1),0)                                                                                   AS SECAS
                                        ,Round(SUM(vacasordeña)/(@fechaFin - @fechaInicio + 1),0)                                                                                  AS PRODUCCION
                                        ,Round(SUM((vqreto + vcreto))/(@fechaFin - @fechaInicio + 1),0)                                                                            AS RETO
                                        ,Round(SUM((jaulas + destetadas + destetadas2 + vaquillas))/(@fechaFin - @fechaInicio + 1),0) AS Crianza
                                        ,Round(SUM((jaulas + destetadas + destetadas2 + vaquillas + vacassecas + vacasordeña + vqreto + vcreto))/(@fechaFin - @fechaInicio + 1),0) AS InventarioTotal
                                FROM INVENTARIOAFIXDIA
                                WHERE fecha BETWEEN @fechaInicio AND @fechaFin";
                query = query.Replace("@fechaInicio", LibreriaGTH.Utils.ConvertToJulian(fechaInicio).ToString()).Replace("@fechaFin", LibreriaGTH.Utils.ConvertToJulian(fechaFin).ToString());

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new InventarioAfiXDia()
                {
                    Ordeño = x["PRODUCCION"] != DBNull.Value ? Convert.ToDecimal(x["PRODUCCION"]) : 0,
                    Secas = x["SECAS"] != DBNull.Value ? Convert.ToDecimal(x["SECAS"]) : 0,
                    Reto = x["RETO"] != DBNull.Value ? Convert.ToDecimal(x["RETO"]) : 0,
                    Jaulas = x["JAULAS_"] != DBNull.Value ? Convert.ToDecimal(x["JAULAS_"]) : 0,
                    Crecimiento = x["CRECIMIENTO"] != DBNull.Value ? Convert.ToDecimal(x["CRECIMIENTO"]) : 0,
                    Desarrollo = x["DESARROLLO"] != DBNull.Value ? Convert.ToDecimal(x["DESARROLLO"]) : 0,
                    Vaquillas = x["VAQUILLAS_"] != DBNull.Value ? Convert.ToDecimal(x["VAQUILLAS_"]) : 0,
                    Crianza = x["Crianza"] != DBNull.Value ? Convert.ToDecimal(x["Crianza"]) : 0,
                    InventarioTotal = x["InventarioTotal"] != DBNull.Value ? Convert.ToDecimal(x["InventarioTotal"]) : 0
                }).ToList().FirstOrDefault();
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "PromedioInventariosAFI", $"Error in {nameof(GetPromedioInventariosAFI)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private decimal GetLecheProducida(DateTime fechaInicio, DateTime fechaFin)
        {
            decimal lecheProducida = 0;
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT IIF( SUM(IIF(m.LECPROD > 0, 1,0)) > 0, CLng(SUM(m.LECPROD) /  SUM(IIF(m.LECPROD > 0, 1,0))), 0) AS LecheProd
                                 FROM MPRODUC m
                                 WHERE m.FECHA BETWEEN @fechaInicio AND @fechaFin";

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@fechaInicio", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                DataTable dt = db.EjecutarConsultaTabla();

                if (dt.Rows.Count > 0)
                    lecheProducida = dt.Rows[0]["LecheProd"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["LecheProd"]) : 0;

            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "LecheProducida", $"Error in {nameof(GetLecheProducida)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return lecheProducida;
        }


        #endregion

        #region Hoja 3
        private List<DatosVacas> GetDatosDesecho(DateTime fechaInicio, DateTime fechaFin, string condicion)
        {
            List<DatosVacas> response = new List<DatosVacas>();
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT  CDATE(FECHA) AS FechaG
                                        ,COUNT(*)     AS Vacas
                                FROM Ddesecho
                                WHERE FECHA BETWEEN @fechaInicio AND @fechaFin 
                                @condicion
                                GROUP BY FECHA  ORDER BY FECHA";
                query = query.Replace("@condicion", condicion);

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@fechaInicio", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new DatosVacas()
                {
                    Fecha = x["FechaG"] != DBNull.Value ? Convert.ToDateTime(x["FechaG"]) : new DateTime(),
                    Vacas = x["Vacas"] != DBNull.Value ? Convert.ToDecimal(x["Vacas"]) : 0
                }).ToList();
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "DatosDesecho", $"Error in {nameof(GetDatosDesecho)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private List<DatosVacas> GetDatosNacimiento(DateTime fechaInicio, DateTime fechaFin, string condicion)
        {
            List<DatosVacas> response = new List<DatosVacas>();
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT  CDATE(FECHA) AS FechaG
                                        ,COUNT(*)     AS Vacas
                                FROM
                                (
	                                SELECT  distinct FECHA
	                                        ,numvac
	                                FROM Dnacimiento
	                                WHERE FECHA BETWEEN @fechaInicio AND @fechaFin
	                                @condicion
                                )
                                GROUP BY  FECHA
                                ORDER BY FECHA";
                query = query.Replace("@condicion", condicion);

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@fechaInicio", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new DatosVacas()
                {
                    Fecha = x["FechaG"] != DBNull.Value ? Convert.ToDateTime(x["FechaG"]) : new DateTime(),
                    Vacas = x["Vacas"] != DBNull.Value ? Convert.ToDecimal(x["Vacas"]) : 0
                }).ToList();
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "DatosNacimiento", $"Error in {nameof(GetDatosNacimiento)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private List<DatosVacas> GetDatosNacimientoMuertas(DateTime fechaInicio, DateTime fechaFin, string condicion)
        {
            List<DatosVacas> response = new List<DatosVacas>();
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT  CDATE(FECHA) AS FechaG
                                        ,COUNT(*)     AS Vacas
                                FROM Dnacimiento
                                WHERE FECHA BETWEEN @fechaInicio AND @fechaFin
                                @condicion
                                GROUP BY  FECHA
                                ORDER BY FECHA";
                query = query.Replace("@condicion", condicion);

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@fechaInicio", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new DatosVacas()
                {
                    Fecha = x["FechaG"] != DBNull.Value ? Convert.ToDateTime(x["FechaG"]) : new DateTime(),
                    Vacas = x["Vacas"] != DBNull.Value ? Convert.ToDecimal(x["Vacas"]) : 0
                }).ToList();

            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "DatosNacimientoMuertas", $"Error in {nameof(GetDatosNacimientoMuertas)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private DatosVacas GetTotalDatosDesecho(DateTime fechaInicio, DateTime fechaFin, string condicion)
        {
            DatosVacas response = new DatosVacas();
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT SUM(Vacas) AS Vacas_
                                FROM (
                                    SELECT  CDATE(FECHA) AS FechaG
                                            ,COUNT(*)     AS Vacas
                                    FROM Ddesecho
                                    WHERE FECHA BETWEEN @fechaInicio AND @fechaFin 
                                    @condicion
                                    GROUP BY FECHA 
                                )T";
                query = query.Replace("@condicion", condicion);

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@fechaInicio", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                DataTable dt = db.EjecutarConsultaTabla();

                if (dt.Rows.Count > 0)
                    response.Vacas = dt.Rows[0]["Vacas_"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["Vacas_"]) : 0;
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "TotalDatosDesecho", $"Error in {nameof(GetTotalDatosDesecho)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private DatosVacas GetTotalDatosNacimiento(DateTime fechaInicio, DateTime fechaFin, string condicion)
        {
            DatosVacas response = new DatosVacas();
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT SUM(Vacas) AS Vacas_
                                FROM (
                                    SELECT  CDATE(FECHA) AS FechaG
                                            ,COUNT(*)     AS Vacas
                                    FROM
                                    (
	                                    SELECT  distinct FECHA
	                                            ,numvac
	                                    FROM Dnacimiento
	                                    WHERE FECHA BETWEEN @fechaInicio AND @fechaFin
	                                    @condicion
                                    )
                                    GROUP BY  FECHA                                
                                )T";
                query = query.Replace("@condicion", condicion);

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@fechaInicio", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                DataTable dt = db.EjecutarConsultaTabla();

                if (dt.Rows.Count > 0)
                    response.Vacas = dt.Rows[0]["Vacas_"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["Vacas_"]) : 0;
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "TotalDatosNacimiento", $"Error in {nameof(GetTotalDatosNacimiento)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private DatosVacas GetTotalDatosNacimientoMuertas(DateTime fechaInicio, DateTime fechaFin, string condicion)
        {
            DatosVacas response = new DatosVacas();
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT SUM(Vacas) AS Vacas_
                                FROM (
                                        SELECT  CDATE(FECHA) AS FechaG
                                                ,COUNT(*)     AS Vacas
                                        FROM Dnacimiento
                                        WHERE FECHA BETWEEN @fechaInicio AND @fechaFin
                                        @condicion
                                        GROUP BY  FECHA
                                )T";
                query = query.Replace("@condicion", condicion);

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@fechaInicio", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                DataTable dt = db.EjecutarConsultaTabla();

                if (dt.Rows.Count > 0)
                    response.Vacas = dt.Rows[0]["Vacas_"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["Vacas_"]) : 0;

            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "TotalDatosNacimientoMuertas", $"Error in {nameof(GetTotalDatosNacimientoMuertas)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private DatosVacas GetTotalDatosAbortos(DateTime fechaInicio, DateTime fechaFin, int condicion)
        {
            DatosVacas response = new DatosVacas();
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT SUM(Vacas) AS Vacas_
                                FROM (
                                    SELECT CDATE(FECHA) AS FechaG
                                            , COUNT(*) AS Vacas
                                     FROM Dabortos 
                                     WHERE FECHA BETWEEN @fechaInicio AND @fechaFin
                                            AND vacvaq = @tipo
                                     GROUP BY FECHA  
                                )T";

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@fechaInicio", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                db.AsignarParametro("@tipo", condicion);
                DataTable dt = db.EjecutarConsultaTabla();

                if (dt.Rows.Count > 0)
                    response.Vacas = dt.Rows[0]["Vacas_"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["Vacas_"]) : 0;
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "TotalDatosAbortos", $"Error in {nameof(GetTotalDatosAbortos)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        public Hoja3 GetTotalReporteHoja3(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, string dia)
        {
            Hoja3 response = new Hoja3() { Dia = dia };

            try
            {
                #region Obtener Datos
                CalostroYOrdeña totalCalostro = GetPromedioCalostro(fechaInicio, fechaFin);
                DatosVacas totalJaulasFuera = GetTotalDatosDesecho(fechaInicio, fechaFin, @" AND LACTANCIA = 0 AND EDOVAC = 15");
                DatosVacas totalJaulasVivas = GetTotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo <> 1 AND destino <> 2 AND (edovac = 1 OR edovac = 10) ");
                DatosVacas totalJaulasMuertas = GetTotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 1 AND (edovac = 1 OR edovac = 10) ");

                DatosVacas totalDesteteVivas = GetTotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo <> 1 AND destino <> 2 AND (edovac = 2 OR edovac = 7) ");
                DatosVacas totalDesteteMuertas = GetTotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 1 AND (edovac = 2 OR edovac = 7) ");

                DatosVacas totalVaquillasMuertas = GetTotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 1 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ");
                DatosVacas totalVaquillasUrgencia = GetTotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 2 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ");
                DatosVacas totalVaquillasDelgadas = GetTotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 3 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ");
                DatosVacas totalVaquillasRegulares = GetTotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 4 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ");
                DatosVacas totalVaquillasGordas = GetTotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 5 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ");
                DatosVacas totalVaquillasOtros = GetTotalDatosDesecho(fechaInicio, fechaFin, @"  AND vacvaq = 2  AND motivo > 5 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ");

                DatosVacas totalVacasMuertas = GetTotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo = 1 AND destino = 1 AND (edovac > 3) ");
                DatosVacas totalVacasUrgencia = GetTotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo = 2 AND destino = 1 AND (edovac > 3) ");
                DatosVacas totalVacasDelgadas = GetTotalDatosDesecho(fechaInicio, fechaFin, @"  AND vacvaq = 1  AND motivo = 3 AND destino = 1 AND (edovac > 3) ");
                DatosVacas totalVacasRegulares = GetTotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo = 4 AND destino = 1 AND (edovac > 3) ");
                DatosVacas totalVacasGordas = GetTotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo = 5 AND destino = 1 AND (edovac > 3) ");
                DatosVacas totalVacasOtros = GetTotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo > 5 AND destino = 1 AND (edovac > 3) ");

                DatosVacas totalPartosVaquillasND = GetTotalDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2  AND (tipopar = 1 or tipopar = 3) ");
                DatosVacas totalPartosVaquillasAbortos = GetTotalDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2  AND (tipopar = 2) ");
                DatosVacas totalPartosVaquillasHembra = GetTotalDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2  AND criaviva = 1  AND criasexo = 1 ");
                DatosVacas totalPartosVaquillasMacho = GetTotalDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2  AND criaviva = 1  AND criasexo = 2 ");

                DatosVacas totalPartosVacasND = GetTotalDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 1  AND (tipopar = 1 or tipopar = 3) ");
                DatosVacas totalPartosVacasAbortos = GetTotalDatosNacimiento(fechaInicio, fechaFin, @"  AND vacvaq = 1  AND (tipopar = 2) ");
                DatosVacas totalPartosVacasHembra = GetTotalDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 1  AND criaviva = 1  AND criasexo = 1 ");
                DatosVacas totalPartosVacasMacho = GetTotalDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 1  AND criaviva = 1  AND criasexo = 2 ");

                DatosVacas totalPartosSinCria = GetTotalDatosNacimiento(fechaInicio, fechaFin, @" AND tipopar <> 2 AND criaviva = 2  AND criasexo = 3 ");
                DatosVacas totalPartosVacas = GetTotalDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 1 ");
                DatosVacas totalPartosVaquillas = GetTotalDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2 ");

                DatosVacas totalMuertasDia = GetTotalDatosNacimientoMuertas(fechaInicio, fechaFin, @" AND criaviva = 2 AND (tipopar = 1 or tipopar = 3) AND criasexo < 3 AND dianoche = 1 ");
                DatosVacas totalMuertasNoche = GetTotalDatosNacimientoMuertas(fechaInicio, fechaFin, @" AND criaviva = 2 AND (tipopar = 1 or tipopar = 3) AND criasexo < 3 AND dianoche = 2 ");

                DatosVacas totalAbortosVacas = GetTotalDatosAbortos(fechaInicio, fechaFin, 1);
                DatosVacas totalAbortosVaquillas = GetTotalDatosAbortos(fechaInicio, fechaFin, 2);
                #endregion

                #region asignar datos
                response.Bp_Gem = totalJaulasFuera != null ? totalJaulasFuera.Vacas : 0;
                response.Jaulas_Vivas = totalJaulasVivas != null ? totalJaulasVivas.Vacas : 0;
                response.Jaulas_Muertas = totalJaulasMuertas != null ? totalJaulasMuertas.Vacas : 0;
                response.Destete_Vivas = totalDesteteVivas != null ? totalDesteteVivas.Vacas : 0;
                response.Destete_Muertas = totalDesteteMuertas != null ? totalDesteteMuertas.Vacas : 0;
                response.Vaquillas_Muertas = totalVaquillasMuertas != null ? totalVaquillasMuertas.Vacas : 0;
                response.Vaquillas_Urgente = totalVaquillasUrgencia != null ? totalVaquillasUrgencia.Vacas : 0;
                response.Vaquillas_RF = totalVaquillasDelgadas != null ? totalVaquillasDelgadas.Vacas : 0;
                response.Vaquillas_RR = totalVaquillasRegulares != null ? totalVaquillasRegulares.Vacas : 0;
                response.Vaquillas_RG = totalVaquillasGordas != null ? totalVaquillasGordas.Vacas : 0;
                response.Vaquillas_Otros = totalVaquillasOtros != null ? totalVaquillasOtros.Vacas : 0;
                response.Vacas_Muertas = totalVacasMuertas != null ? totalVacasMuertas.Vacas : 0;
                response.Vacas_Urgente = totalVacasUrgencia != null ? totalVacasUrgencia.Vacas : 0;
                response.Vacas_RF = totalVacasDelgadas != null ? totalVacasDelgadas.Vacas : 0;
                response.Vacas_RR = totalVacasRegulares != null ? totalVacasRegulares.Vacas : 0;
                response.Vacas_RG = totalVacasGordas != null ? totalVacasGordas.Vacas : 0;
                response.Vacas_Otros = totalVacasOtros != null ? totalVacasOtros.Vacas : 0;
                response.Vaquillas_ND = totalPartosVaquillasND != null ? totalPartosVaquillasND.Vacas : 0;
                response.Vaquillas_A = totalPartosVaquillasAbortos != null ? totalPartosVaquillasAbortos.Vacas : 0;
                response.Vaquillas_Vivos_H = totalPartosVaquillasHembra != null ? totalPartosVaquillasHembra.Vacas : 0;
                response.Vaquillas_Vivos_M = totalPartosVaquillasMacho != null ? totalPartosVaquillasMacho.Vacas : 0;
                response.Vacas_ND = totalPartosVacasND != null ? totalPartosVacasND.Vacas : 0;
                response.Vacas_A = totalPartosVacasAbortos != null ? totalPartosVacasAbortos.Vacas : 0;
                response.Vacas_Vivos_H = totalPartosVacasHembra != null ? totalPartosVacasHembra.Vacas : 0;
                response.Vacas_Vivos_M = totalPartosVacasMacho != null ? totalPartosVacasMacho.Vacas : 0;
                response.SC = totalPartosSinCria != null ? totalPartosSinCria.Vacas : 0;
                response.Partos_Vacas = totalPartosVacas != null ? totalPartosVacas.Vacas : 0;
                response.Partos_Vaquillas = totalPartosVaquillas != null ? totalPartosVaquillas.Vacas : 0;
                response.Muertas_Dia = totalMuertasDia != null ? totalMuertasDia.Vacas : 0;
                response.Muertas_Noc = totalMuertasNoche != null ? totalMuertasNoche.Vacas : 0;

                response.Diferencia_Calostro = totalCalostro != null ? totalCalostro.Diferencia : 0;
                response.Porcentaje_Calostro = totalCalostro != null ? totalCalostro.Porcentaje : 0;
                response.Calidad_Calostro = totalCalostro.Calidad != null ? totalCalostro.Calidad : 0;
                #endregion
            }
            catch { }


            return response;
        }
        #endregion

        #region Hoja 4
        private List<DatosVacas> GetDatosSalud(DateTime fechaInicio, DateTime fechaFin, string condicion)
        {
            List<DatosVacas> response = new List<DatosVacas>();
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT  CDATE(FECHA) AS FechaG
                                        ,COUNT(*)  AS Vacas
                                FROM DSALUD 
                                WHERE FECHA BETWEEN  @fechaInicio AND @fechaFin
                                @condicion
                                GROUP BY  FECHA
                                ORDER BY FECHA";
                query = query.Replace("@condicion", condicion);

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@fechaInicio", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new DatosVacas()
                {
                    Fecha = x["FechaG"] != DBNull.Value ? Convert.ToDateTime(x["FechaG"]) : new DateTime(),
                    Vacas = x["Vacas"] != DBNull.Value ? Convert.ToDecimal(x["Vacas"]) : 0
                }).ToList();

            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "DatosSalud", $"Error in {nameof(GetDatosSalud)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private List<ValInventario> GetDatosValInventario(DateTime fechaInicio, DateTime fechaFin)
        {
            List<ValInventario> response = new List<ValInventario>();
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT  CDATE(FECHA) AS FechaG
                                        ,preñvc       AS VacasPreñadas
                                        ,vaciavc      AS VacasVacias
                                        ,preñvq       AS VaquillasPreñadas
                                        ,vaciavq      AS VaquillasVacias
                                FROM VALINVENTARIO
                                WHERE FECHA BETWEEN @fechaInicio AND @fechaFin
                                ORDER BY FECHA";

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@fechaInicio", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new ValInventario()
                {
                    Fecha = x["FechaG"] != DBNull.Value ? Convert.ToDateTime(x["FechaG"]) : new DateTime(),
                    VacasPreñadas = x["VacasPreñadas"] != DBNull.Value ? Convert.ToDecimal(x["VacasPreñadas"]) : 0,
                    VacasVacias = x["VacasVacias"] != DBNull.Value ? Convert.ToDecimal(x["VacasVacias"]) : 0,
                    VaquillasPreñadas = x["VaquillasPreñadas"] != DBNull.Value ? Convert.ToDecimal(x["VaquillasPreñadas"]) : 0,
                    VaquillasVacias = x["VaquillasVacias"] != DBNull.Value ? Convert.ToDecimal(x["VaquillasVacias"]) : 0

                }).ToList();

            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "DatosValInventario", $"Error in {nameof(GetDatosValInventario)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private List<DatosVacas> GetDatosAbortos(DateTime fechaInicio, DateTime fechaFin, int condicion)
        {
            List<DatosVacas> response = new List<DatosVacas>();
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT CDATE(FECHA) AS FechaG
                                        , COUNT(*) AS Vacas
                                 FROM Dabortos 
                                 WHERE FECHA BETWEEN @fechaInicio AND @fechaFin
                                        AND vacvaq = @tipo
                                 GROUP BY FECHA  ORDER BY FECHA";

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@fechaInicio", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                db.AsignarParametro("@tipo", condicion);
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new DatosVacas()
                {
                    Fecha = x["FechaG"] != DBNull.Value ? Convert.ToDateTime(x["FechaG"]) : new DateTime(),
                    Vacas = x["Vacas"] != DBNull.Value ? Convert.ToDecimal(x["Vacas"]) : 0
                }).ToList();
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "DatosAbortos", $"Error in {nameof(GetDatosAbortos)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private DatosVacas GetTotalDatosSalud(DateTime fechaInicio, DateTime fechaFin, string condicion)
        {
            DatosVacas response = new DatosVacas();
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT SUM(Vacas) AS Vacas_
                                FROM(
                                    SELECT  CDATE(FECHA) AS FechaG
                                            ,COUNT(*)  AS Vacas
                                    FROM DSALUD 
                                    WHERE FECHA BETWEEN  @fechaInicio AND @fechaFin
                                    @condicion
                                    GROUP BY  FECHA
                                    ORDER BY FECHA
                                )T";
                query = query.Replace("@condicion", condicion);

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@fechaInicio", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                DataTable dt = db.EjecutarConsultaTabla();

                if (dt.Rows.Count > 0)
                    response.Vacas = dt.Rows[0]["Vacas_"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["Vacas_"]) : 0;

            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "TotalDatosSalud", $"Error in {nameof(GetTotalDatosSalud)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private ValInventario GetTotalDatosValInventario(DateTime fechaInicio, DateTime fechaFin)
        {
            ValInventario response = new ValInventario();
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT  SUM(preñvc)       AS VacasPreñadas
                                        ,SUM(vaciavc)      AS VacasVacias
                                        ,SUM(preñvq)       AS VaquillasPreñadas
                                        ,SUM(vaciavq)      AS VaquillasVacias
                                FROM VALINVENTARIO
                                WHERE FECHA BETWEEN @fechaInicio AND @fechaFin";

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                db.AsignarParametro("@fechaInicio", LibreriaGTH.Utils.ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", LibreriaGTH.Utils.ConvertToJulian(fechaFin));
                DataTable dt = db.EjecutarConsultaTabla();

                if (dt.Rows.Count > 0)
                {
                    response.VacasPreñadas = dt.Rows[0]["VacasPreñadas"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["VacasPreñadas"]) : 0;
                    response.VacasVacias = dt.Rows[0]["VacasVacias"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["VacasVacias"]) : 0;
                    response.VaquillasPreñadas = dt.Rows[0]["VaquillasPreñadas"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["VaquillasPreñadas"]) : 0;
                    response.VaquillasVacias = dt.Rows[0]["VaquillasVacias"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["VaquillasVacias"]) : 0;
                }
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "TotalDatosValInventario", $"Error in {nameof(GetTotalDatosValInventario)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        public Hoja4 GetTotalReporteHoja4(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, string dia)
        {
            Hoja4 response = new Hoja4();
            response.Dia = dia;

            try
            {
                DatosVacas datosUbre_MA = GetTotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 1 AND enfdetalle = 1 ");
                DatosVacas datosUbre_SL = GetTotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 1 AND enfdetalle = 2 ");
                DatosVacas datosMetabolicos_FL = GetTotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 5 AND enfdetalle = 1 ");
                DatosVacas datosMetabolicos_CET = GetTotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 5 AND enfdetalle = 2 ");
                DatosVacas datosLocomotores_BE = GetTotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 4 AND enfdetalle = 1 ");
                DatosVacas datosLocomotores_TRA = GetTotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 4 AND enfdetalle = 2 ");
                DatosVacas datosLocomotores_GA = GetTotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 4 AND enfdetalle = 3 ");
                DatosVacas datosDigestivos_AC = GetTotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 1 ");
                DatosVacas datosDigestivos_ES = GetTotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 2 ");
                DatosVacas datosDigestivos_DI = GetTotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 3 ");
                DatosVacas datosDigestivos_TI = GetTotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 4 ");
                DatosVacas datosReproductivos_RE = GetTotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 1 ");
                DatosVacas datosReproductivos_ME = GetTotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 2 ");
                DatosVacas datosReproductivos_PIO = GetTotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 3 ");
                DatosVacas datosReproductivos_QUI = GetTotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 4 ");
                DatosVacas datosReproductivos_CS = GetTotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 5 ");
                DatosVacas datosRespiratorios_Neu = GetTotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 7 AND enfdetalle = 1 ");
                DatosVacas datosBecerras_Neu = GetTotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 2  AND enfermedad = 8 AND enfdetalle = 1 ");
                DatosVacas datosBecerras_Fie = GetTotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 2  AND enfermedad = 8 AND enfdetalle = 2 ");
                DatosVacas datosBecerras_Di = GetTotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 2  AND enfermedad = 8 AND enfdetalle = 3 ");
                DatosVacas datosBecerras_Conj = GetTotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 2  AND enfermedad = 8 AND enfdetalle = 4 ");
                ValInventario datosValInventarios = GetTotalDatosValInventario(fechaInicio, fechaFin);
                DatosVacas datosAbortosVacas = GetTotalDatosAbortos(fechaInicio, fechaFin, 1);
                DatosVacas datosAbortosVaquillas = GetTotalDatosAbortos(fechaInicio, fechaFin, 2);

                response.Ubre_MA = datosUbre_MA.Vacas;
                response.Ubre_SL = datosUbre_SL.Vacas;
                response.Metabolicos_FL = datosMetabolicos_FL.Vacas;
                response.Metabolicos_CET = datosMetabolicos_CET.Vacas;
                response.Locomotores_BE = datosLocomotores_BE.Vacas;
                response.Locomotores_TRA = datosLocomotores_TRA.Vacas;
                response.Locomotores_GA = datosLocomotores_GA.Vacas;
                response.Digestivos_AC = datosDigestivos_AC.Vacas;
                response.Digestivos_ES = datosDigestivos_ES.Vacas;
                response.Digestivos_DI = datosDigestivos_DI.Vacas;
                response.Digestivos_TI = datosDigestivos_TI.Vacas;
                response.Reproductivos_RE = datosReproductivos_RE.Vacas;
                response.Reproductivos_ME = datosReproductivos_ME.Vacas;
                response.Reproductivos_PIO = datosReproductivos_PIO.Vacas;
                response.Reproductivos_QUI = datosReproductivos_QUI.Vacas;
                response.Reproductivos_CS = datosReproductivos_CS.Vacas;
                response.Respiratorios_Neu = datosRespiratorios_Neu.Vacas;
                response.Becerras_Neu = datosBecerras_Neu.Vacas;
                response.Becerras_Fie = datosBecerras_Fie.Vacas;
                response.Becerras_Di = datosBecerras_Di.Vacas;
                response.Becerras_Conj = datosBecerras_Conj.Vacas;

                response.Vacas_Pren = datosValInventarios.VacasPreñadas;
                response.Vacas_Vacias = datosValInventarios.VacasVacias;
                response.Vacas_Diag = response.Vacas_Pren + response.Vacas_Vacias;
                response.Vacas_Porcentaje_Pren = response.Vacas_Diag > 0 ? (response.Vacas_Pren / response.Vacas_Diag) * 100 : 0;
                response.Vacas_Porcentaje_Vacias = response.Vacas_Diag > 0 ? (response.Vacas_Vacias / response.Vacas_Diag) * 100 : 0;

                response.Vaquillas_Pren = datosValInventarios.VaquillasPreñadas;
                response.Vaquillas_Vacias = datosValInventarios.VaquillasVacias;
                response.Vaquillas_Diag = response.Vaquillas_Pren + response.Vaquillas_Vacias;
                response.Vaquillas_Porcentaje_Pren = response.Vaquillas_Diag > 0 ? (response.Vaquillas_Pren / response.Vaquillas_Diag) * 100 : 0;
                response.Vaquillas_Porcentaje_Vacias = response.Vaquillas_Diag > 0 ? (response.Vaquillas_Vacias / response.Vaquillas_Diag) * 100 : 0;

                response.Abortos_Vaquillas = datosAbortosVaquillas.Vacas;
                response.Abortos_Vacas = datosAbortosVacas.Vacas;
            }
            catch { }

            return response;
        }
        #endregion

        #region log
        public void InsLog(string interfaceCode, string systemSource, string keyValue, string errorMessage,
          string interfaceMessage, string exceptionMessage, Guid guid)
        {
            ModeloDatosSQL db = new ModeloDatosSQL(_sqlConnection);

            try
            {
                db.Conectar();
                db.CrearComando("usp_ins_InterfaceLogTbl", tipoComandoSQL.storedProcedure);
                db.AsignarParametro("@p_InterfaceCode", interfaceCode);
                db.AsignarParametro("@p_SystemSource", systemSource);
                db.AsignarParametro("@p_KeyValue", keyValue);
                db.AsignarParametro("@p_ErrorMessage", errorMessage);
                db.AsignarParametro("@p_InterfaceMessage", interfaceMessage);
                db.AsignarParametro("@p_ExceptionMessage", exceptionMessage);
                db.AsignarParametro("@p_MessageId", guid);
                db.EjecutarComando();
            }
            catch (DbException ex)
            {
                Console.WriteLine("Error in method InsLog:" + ex.Message);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }
        }
        #endregion

    }
}
