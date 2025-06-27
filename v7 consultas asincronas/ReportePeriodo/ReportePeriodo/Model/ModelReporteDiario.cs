using ReportePeriodo.Data;
using ReportePeriodo.Entity;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using gth = LibreriaGTH.Consumos.Modelo;
using gthEntity = LibreriaGTH.Consumos.Entidad;

namespace ReportePeriodo.Model
{
    public class ModelReporteDiario
    {
        #region variables
        List<DateTime> fechasReporte;

        #region variables de conexion de base de datos
        string _sqlConnection;
        string _oleDbConnection;
        string _fdbConnection;
        #endregion

        gth.ConsumoModel consumoModel;

        #region indicadores Teoricos Alimentacion
        List<gthEntity.IndicadorReporteTeorico> indicadorProduccion;
        List<gthEntity.IndicadorReporteTeorico> indicadorSecas;
        List<gthEntity.IndicadorReporteTeorico> indicadorReto;
        List<gthEntity.IndicadorReporteTeorico> indicadorCrecimiento;
        List<gthEntity.IndicadorReporteTeorico> indicadorDesarrollo;
        List<gthEntity.IndicadorReporteTeorico> indicadorVaquillas;
        #endregion

        #region indicadores Alimentacion
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
        #endregion

        #region etapas Alimentacion
        gthEntity.EtapaTyp etapaProduccion = new gthEntity.EtapaTyp() { Id = 1, Clave = "10", Nombre = "Produccion" };
        gthEntity.EtapaTyp etapaSecas = new gthEntity.EtapaTyp() { Id = 2, Clave = "21", Nombre = "Secas" };
        gthEntity.EtapaTyp etapaReto = new gthEntity.EtapaTyp() { Id = 3, Clave = "22", Nombre = "Reto" };
        gthEntity.EtapaTyp etapaJaulas = new gthEntity.EtapaTyp() { Id = 4, Clave = "31", Nombre = "Jaulas" };
        gthEntity.EtapaTyp etapaCrecimiento = new gthEntity.EtapaTyp() { Id = 5, Clave = "32", Nombre = "Crecimiento" };
        gthEntity.EtapaTyp etapaDesarrollo = new gthEntity.EtapaTyp() { Id = 6, Clave = "33", Nombre = "Desarrollo" };
        gthEntity.EtapaTyp etapaVaquillas = new gthEntity.EtapaTyp() { Id = 7, Clave = "34", Nombre = "Vaquillas Preñadas" };
        #endregion

        #region Listas Hoja 1 y 2
        List<DatosProduccion> _datosDeProduccion;
        List<NivelCribas> _datosNivelDeCribas;
        List<Inventario> _datosInventario;
        List<CalostroYOrdeña> _datosCalostro;
        List<Mproduc> _datosMproduc;
        List<InventarioAfiXDia> _datosInventarioAFI;
        List<Mproduc2> _datosMproduc2;
        List<BacteriologiaLeche> _datosBacteriologia;
        #endregion

        #region variables para promedios
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

        #region Listas Hoja 3
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
        #endregion

        #region Listas Hoja 4
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
        #endregion

        #endregion


        #region constructor
        public ModelReporteDiario(string oleDbConnection, string sqlConnection, string fdbConnection)
        {
            this._oleDbConnection = oleDbConnection;
            this._sqlConnection = sqlConnection;
            this._fdbConnection = fdbConnection;

            consumoModel = new gth.ConsumoModel(sqlConnection);
            fechasReporte = new List<DateTime>();

            #region indicadores Teoricos Alimentacion
            indicadorProduccion = new List<gthEntity.IndicadorReporteTeorico>();
            indicadorSecas = new List<gthEntity.IndicadorReporteTeorico>();
            indicadorReto = new List<gthEntity.IndicadorReporteTeorico>();
            indicadorCrecimiento = new List<gthEntity.IndicadorReporteTeorico>();
            indicadorDesarrollo = new List<gthEntity.IndicadorReporteTeorico>();
            indicadorVaquillas = new List<gthEntity.IndicadorReporteTeorico>();
            #endregion

            #region indicadores Alimentacion
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
            #endregion

            #region Listas Hoja 1 y 2
            _datosDeProduccion = new List<DatosProduccion>();
            _datosNivelDeCribas = new List<NivelCribas>();
            _datosInventario = new List<Inventario>();
            _datosCalostro = new List<CalostroYOrdeña>();
            _datosMproduc = new List<Mproduc>();
            _datosInventarioAFI = new List<InventarioAfiXDia>();
            _datosMproduc2 = new List<Mproduc2>();
            _datosBacteriologia = new List<BacteriologiaLeche>();
            #endregion


        }
        #endregion

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
            get { return _totalHoja1; }
        }

        public Hoja1 TotalHoja1AñoAnt
        {
            get { return _totalHoja1AñoAnt; }
        }

        public List<CalostroYOrdeña> DatosCalostro
        {
            get { return _datosCalostro; }
        }

        public List<BacteriologiaLeche> DatosBacteriologia
        {
            get { return _datosBacteriologia; }
        }

        public Hoja2 TotalHoja2
        {
            get { return _totalHoja2; }
        }

        public Hoja2 TotalHoja2AñoAnt
        {
            get { return _totalHoja2AñoAnt; }
        }

        public CalostroYOrdeña PromedioCalostroOrdeño
        {
            get { return _promedioCalostro; }
        }

        public CalostroYOrdeña PromedioCalostroAñoAnt
        {
            get { return _promedioCalostroAñoAnt; }
        }

        public decimal PrecioLeche
        {
            get
            {
                return _precioLeche;
            }
        }
        public decimal PrecioLecheAñoAnt
        {
            get
            {
                return _precioLecheAñoAnt;
            }
        }

        #endregion

        #region metodos publicos
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

        public void CargarDatosAlimentacion(gthEntity.CentroTyp centro, DateTime fechaInicio, DateTime fechaFin)
        {
            List<gthEntity.CentroTyp> centroAux = new List<gthEntity.CentroTyp>();
            centroAux.Add(centro);

            var produccionTask = new Task(() => { IndicadorProduccion(centroAux, fechaInicio, fechaFin); });
            var secasTask = new Task(() => { IndicadorSecas(centroAux, fechaInicio, fechaFin); });
            var retoTask = new Task(() => { IndicadorReto(centroAux, fechaInicio, fechaFin); });
            var jaulasTask = new Task(() => { IndicadorJaulas(centroAux, fechaInicio, fechaFin); });
            var crecimientoTask = new Task(() => { IndicadorCrecimiento(centroAux, fechaInicio, fechaFin); });
            var desarrolloTask = new Task(() => { IndicadorDesarrollo(centroAux, fechaInicio, fechaFin); });
            var vaquillasTask = new Task(() => { IndicadorVaquillas(centroAux, fechaInicio, fechaFin); });

            var produccionTeoricoTask = new Task(() => { IndicadoresTeoricoProduccion(centroAux, fechaInicio, fechaFin); });
            var secasTeoricoTask = new Task(() => { IndicadoresTeoricoSecas(centroAux, fechaInicio, fechaFin); });
            var retoTeoricoTask = new Task(() => { IndicadoresTeoricoReto(centroAux, fechaInicio, fechaFin); });
            var crecimientoTeoricoTask = new Task(() => { IndicadoresTeoricoCrecimiento(centroAux, fechaInicio, fechaFin); });
            var desarrolloTeoricoTask = new Task(() => { IndicadoresTeoricoDesarrollo(centroAux, fechaInicio, fechaFin); });
            var vaquillasTeoricoTask = new Task(() => { IndicadoresTeoricoVaquillas(centroAux, fechaInicio, fechaFin); });

            produccionTask.Start();
            secasTask.Start();
            retoTask.Start();
            jaulasTask.Start();
            crecimientoTask.Start();
            desarrolloTask.Start();
            vaquillasTask.Start();

            produccionTeoricoTask.Start();
            secasTeoricoTask.Start();
            retoTeoricoTask.Start();
            crecimientoTeoricoTask.Start();
            desarrolloTeoricoTask.Start();
            vaquillasTeoricoTask.Start();

            var info = new List<Task>
            {
                produccionTask,
                secasTask,
                retoTask,
                jaulasTask,
                crecimientoTask,
                desarrolloTask,
                vaquillasTask,
                produccionTeoricoTask,
                secasTeoricoTask,
                retoTeoricoTask,
                crecimientoTeoricoTask,
                desarrolloTeoricoTask,
                vaquillasTeoricoTask
            };

            Task.WaitAll(info.ToArray());
        }

        public void CargarInformacionHojas(gthEntity.CentroTyp centro, Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            var hoja1Task = new Task(() => { CargaInformacionHoja1(centro, rancho, fechaInicio, fechaFin); });
            var hoja3Task = new Task(() => { CargaInformacionHoja3(centro, rancho, fechaInicio, fechaFin); });
            var hoja4Task = new Task(() => { CargaInformacionHoja4(centro, rancho, fechaInicio, fechaFin); });
            var mproducTask = new Task(() => { PromedioMproduc(rancho, fechaInicio, fechaFin); });
            var inventarioTask = new Task(() => { TotalInventarioAFI(fechaInicio, fechaFin); });
            var calostroTask = new Task(() => { PromedioCalostro(fechaInicio, fechaFin); });
            var dpTask = new Task(() => { TotalDatosDeProduccion(fechaInicio, fechaFin); });
            var lecheTask = new Task(() => { PrecioLecheTask(centro, fechaInicio, fechaFin); });
            var leche2Task = new Task(() => { LecheProducidaTask(centro, fechaInicio, fechaFin); });

            hoja1Task.Start();
            hoja3Task.Start();
            hoja4Task.Start();
            mproducTask.Start();
            inventarioTask.Start();
            calostroTask.Start();
            dpTask.Start();
            lecheTask.Start();
            leche2Task.Start();

            var info = new List<Task>
            {
                hoja1Task,
                hoja3Task,
                hoja4Task,
                mproducTask,
                inventarioTask,
                calostroTask,
                dpTask,
                lecheTask,
                leche2Task
            };

            Task.WaitAll(info.ToArray());
        }

        public List<Hoja1> GetReporteHoja1(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, out List<decimal?> listaDecimal)
        {
            List<Hoja1> response = new List<Hoja1>();
            listaDecimal = new List<decimal?>();

            try
            {
                List<DateTime> fechas = GetFechas(fechaFin);
                DatosProduccion busquedaDatosProduccion;
                Mproduc busquedaMProduc;
                NivelCribas busquedaCribas;
                CalostroYOrdeña busquedaCalostro;
                Inventario busquedaInventario;

                foreach (DateTime fecha in fechas)
                {
                    Hoja1 hoja = new Hoja1();
                    hoja.Dia = fecha.Day.ToString();

                    busquedaDatosProduccion = (from x in _datosDeProduccion where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    busquedaMProduc = (from x in _datosMproduc where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    busquedaCribas = (from x in _datosNivelDeCribas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    busquedaCalostro = (from x in _datosCalostro where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    busquedaInventario = (from x in _datosInventario where x.Fecha == fecha select x).ToList().FirstOrDefault();

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

                List<DatosProduccion> busquedaDec1 = (from x in _datosDeProduccion where x.Fecha.Day < 11 select x).ToList();
                List<DatosProduccion> busquedaDec1Dif0 = (from x in _datosDeProduccion where x.Fecha.Day < 11 && x.Grasa != 0 select x).ToList();

                decimal? decPorcGrasa = 0; decimal? valordecPorcGrasa = 0;
                foreach (DatosProduccion itemDec in busquedaDec1)
                    decPorcGrasa += itemDec.Grasa;

                valordecPorcGrasa = busquedaDec1Dif0.Count > 0 ? decPorcGrasa / busquedaDec1Dif0.Count : 0;
                listaDecimal.Add(valordecPorcGrasa);

                if (fechaFin.Day > 10)
                {
                    decPorcGrasa = 0; valordecPorcGrasa = 0;
                    List<DatosProduccion> busquedaDec2 = (from x in _datosDeProduccion where x.Fecha.Day > 10 && x.Fecha.Day < 21 select x).ToList();
                    List<DatosProduccion> busquedaDec2Dif0 = (from x in _datosDeProduccion where x.Fecha.Day > 10 && x.Fecha.Day < 21 && x.Grasa != 0 select x).ToList();

                    foreach (DatosProduccion itemDec in busquedaDec2)
                    {
                        decPorcGrasa += itemDec.Grasa;
                    }
                    valordecPorcGrasa = busquedaDec2Dif0.Count > 0 ? decPorcGrasa / busquedaDec2Dif0.Count : 0;
                    listaDecimal.Add(valordecPorcGrasa);
                }
                else
                {
                    listaDecimal.Add(0);
                }

                if (fechaFin.Day > 20)
                {
                    decPorcGrasa = 0; valordecPorcGrasa = 0;
                    List<DatosProduccion> busquedaDec3 = (from x in _datosDeProduccion where x.Fecha.Day > 21 select x).ToList();
                    List<DatosProduccion> busquedaDec3Dif0 = (from x in _datosDeProduccion where x.Fecha.Day > 21 && x.Grasa != 0 select x).ToList();

                    foreach (DatosProduccion itemDec in busquedaDec3)
                    {
                        decPorcGrasa += itemDec.Grasa;
                    }
                    valordecPorcGrasa = busquedaDec3Dif0.Count > 0 ? decPorcGrasa / busquedaDec3Dif0.Count : 0;
                    listaDecimal.Add(valordecPorcGrasa);
                }
                else
                {
                    listaDecimal.Add(0);
                }
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "GetReporteHoja1", $"Error in {nameof(GetReporteHoja1)}", ex.Message, ex.Message, ex.StackTrace, guid);
            }

            return response;
        }

        public void GenerarPromedios(Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            var promediosHoja1 = new Task(() => { PromediosHoja1(rancho); });
            var promediosHoja2 = new Task(() => { PromediosHoja2(); });
            var promediosHoja3 = new Task(() => { PromediosHoja3(fechaInicio, fechaFin); });
            var promediosHoja4 = new Task(() => { PromediosHoja4(fechaInicio, fechaFin); });

            promediosHoja1.Start();
            promediosHoja2.Start();
            promediosHoja3.Start();
            promediosHoja4.Start();

            var infoTask = new List<Task>
            {
                promediosHoja1,
                promediosHoja2,
                promediosHoja3,
                promediosHoja4
            };

            Task.WaitAll(infoTask.ToArray());
        }

        public List<Hoja2> GetReporteHoja2(Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            List<Hoja2> response = new List<Hoja2>();

            try
            {
                List<DateTime> fechas = GetFechas(fechaFin);
                InventarioAfiXDia busqudaInventarioAFI;
                Mproduc2 busquedaMproduc;

                foreach (DateTime fecha in fechas)
                {
                    busqudaInventarioAFI = (from x in _datosInventarioAFI where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    busquedaMproduc = (from x in _datosMproduc2 where x.Fecha == fecha select x).ToList().FirstOrDefault();

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
                InsLog("ReporteDiario", "GetReporteHoja2", $"Error in {nameof(GetReporteHoja2)}", ex.Message, ex.Message, ex.StackTrace, guid);
            }

            return response;
        }

        public List<Hoja3> GetReporteHoja3(Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            List<Hoja3> response = new List<Hoja3>();

            try
            {
                List<DateTime> fechasReporte = GetFechas(fechaFin);

                foreach (DateTime fecha in fechasReporte)
                {
                    Hoja3 item = new Hoja3();

                    item.Dia = fecha.Day.ToString();

                    #region obtener datos de la lista
                    CalostroYOrdeña busquedaCalostro = (from x in _datosCalostro where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaJaulasFuera = (from x in _datosJaulasFuera where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaJaulasVivas = (from x in _datosJaulasVivas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaJaulasMuertas = (from x in _datosJaulasMuertas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaDesteteVivas = (from x in _datosDesteteVivas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaDesteteMuertas = (from x in _datosDesteteMuertas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaVaquillasMuertas = (from x in _datosVaquillasMuertas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaVaquillasUrgencia = (from x in _datosVaquillasUrgencia where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaVaquillasDelgadas = (from x in _datosVaquillasDelgadas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaVaquillasRegulares = (from x in _datosVaquillasRegulares where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaVaquillasGordas = (from x in _datosVaquillasGordas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaVaquillasOtros = (from x in _datosVaquillasOtros where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaVacasMuertas = (from x in _datosVacasMuertas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaVacasUrgencia = (from x in _datosVacasUrgencia where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaVacasDelgadas = (from x in _datosVacasDelgadas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaVacasRegulares = (from x in _datosVacasRegulares where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaVacasGordas = (from x in _datosVacasGordas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaVacasOtros = (from x in _datosVacasOtros where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaPartosVaquillasND = (from x in _datosPartosVaquillasND where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaPartosVaquillasAbortos = (from x in _datosPartosVaquillasAbortos where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaPartosVaquillasHembra = (from x in _datosPartosVaquillasHembra where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaPartosVaquillasMacho = (from x in _datosPartosVaquillasMacho where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaPartosVacasND = (from x in _datosPartosVacasND where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaPartosVacasAbortos = (from x in _datosPartosVacasAbortos where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaPartosVacasHembra = (from x in _datosPartosVacasHembra where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaPartosVacasMacho = (from x in _datosPartosVacasMacho where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaPartosSinCria = (from x in _datosPartosSinCria where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaPartosVacas = (from x in _datosPartosVacas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaPartosVaquillas = (from x in _datosPartosVaquillas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaMuertasDia = (from x in _datosMuertasDia where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaMuertasNoche = (from x in _datosMuertasNoche where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    #endregion

                    #region asignar datos
                    item.Bp_Gem = busquedaJaulasFuera != null ? busquedaJaulasFuera.Vacas : 0;
                    item.Jaulas_Vivas = busquedaJaulasVivas != null ? busquedaJaulasVivas.Vacas : 0;
                    item.Jaulas_Muertas = busquedaJaulasMuertas != null ? busquedaJaulasMuertas.Vacas : 0;
                    item.Destete_Vivas = busquedaDesteteVivas != null ? busquedaDesteteVivas.Vacas : 0;
                    item.Destete_Muertas = busquedaDesteteMuertas != null ? busquedaDesteteMuertas.Vacas : 0;
                    item.Vaquillas_Muertas = busquedaVaquillasMuertas != null ? busquedaVaquillasMuertas.Vacas : 0;
                    item.Vaquillas_Urgente = busquedaVaquillasUrgencia != null ? busquedaVaquillasUrgencia.Vacas : 0;
                    item.Vaquillas_RF = busquedaVaquillasDelgadas != null ? busquedaVaquillasDelgadas.Vacas : 0;
                    item.Vaquillas_RR = busquedaVaquillasRegulares != null ? busquedaVaquillasRegulares.Vacas : 0;
                    item.Vaquillas_RG = busquedaVaquillasGordas != null ? busquedaVaquillasGordas.Vacas : 0;
                    item.Vaquillas_Otros = busquedaVaquillasOtros != null ? busquedaVaquillasOtros.Vacas : 0;
                    item.Vacas_Muertas = busquedaVacasMuertas != null ? busquedaVacasMuertas.Vacas : 0;
                    item.Vacas_Urgente = busquedaVacasUrgencia != null ? busquedaVacasUrgencia.Vacas : 0;
                    item.Vacas_RF = busquedaVacasDelgadas != null ? busquedaVacasDelgadas.Vacas : 0;
                    item.Vacas_RR = busquedaVacasRegulares != null ? busquedaVacasRegulares.Vacas : 0;
                    item.Vacas_RG = busquedaVacasGordas != null ? busquedaVacasGordas.Vacas : 0;
                    item.Vacas_Otros = busquedaVacasOtros != null ? busquedaVacasOtros.Vacas : 0;
                    item.Vaquillas_ND = busquedaPartosVaquillasND != null ? busquedaPartosVaquillasND.Vacas : 0;
                    item.Vaquillas_A = busquedaPartosVaquillasAbortos != null ? busquedaPartosVaquillasAbortos.Vacas : 0;
                    item.Vaquillas_Vivos_H = busquedaPartosVaquillasHembra != null ? busquedaPartosVaquillasHembra.Vacas : 0;
                    item.Vaquillas_Vivos_M = busquedaPartosVaquillasMacho != null ? busquedaPartosVaquillasMacho.Vacas : 0;
                    item.Vacas_ND = busquedaPartosVacasND != null ? busquedaPartosVacasND.Vacas : 0;
                    item.Vacas_A = busquedaPartosVacasAbortos != null ? busquedaPartosVacasAbortos.Vacas : 0;
                    item.Vacas_Vivos_H = busquedaPartosVacasHembra != null ? busquedaPartosVacasHembra.Vacas : 0;
                    item.Vacas_Vivos_M = busquedaPartosVacasMacho != null ? busquedaPartosVacasMacho.Vacas : 0;
                    item.SC = busquedaPartosSinCria != null ? busquedaPartosSinCria.Vacas : 0;
                    item.Partos_Vacas = busquedaPartosVacas != null ? busquedaPartosVacas.Vacas : 0;
                    item.Partos_Vaquillas = busquedaPartosVaquillas != null ? busquedaPartosVaquillas.Vacas : 0;
                    item.Muertas_Dia = busquedaMuertasDia != null ? busquedaMuertasDia.Vacas : 0;
                    item.Muertas_Noc = busquedaMuertasNoche != null ? busquedaMuertasNoche.Vacas : 0;
                    item.Diferencia_Calostro = busquedaCalostro != null ? busquedaCalostro.Diferencia : 0;
                    item.Porcentaje_Calostro = busquedaCalostro != null ? busquedaCalostro.Porcentaje : 0;
                    item.Calidad_Calostro = busquedaCalostro != null ? busquedaCalostro.Calidad : 0;
                    #endregion

                    response.Add(item);
                }
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "GetReporteHoja3", $"Error in {nameof(GetReporteHoja3)}", ex.Message, ex.Message, ex.StackTrace, guid);
            }

            return response;
        }

        public List<Hoja4> GetReporteHoja4(Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            List<Hoja4> response = new List<Hoja4>();

            try
            {
                List<DateTime> fechasReporte = GetFechas(fechaFin);

                foreach (DateTime fecha in fechasReporte)
                {
                    Hoja4 newItem = new Hoja4();
                    newItem.Dia = fecha.Day.ToString();

                    #region tomar datos por dia
                    DatosVacas busquedaUbre_MA = (from x in _datosUbre_MA where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaUbre_SL = (from x in _datosUbre_SL where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaMetabolicos_FL = (from x in _datosMetabolicos_FL where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaMetabolicos_CET = (from x in _datosMetabolicos_CET where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaLocomotores_BE = (from x in _datosLocomotores_BE where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaLocomotores_TRA = (from x in _datosLocomotores_TRA where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaLocomotores_GA = (from x in _datosLocomotores_GA where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaDigestivos_AC = (from x in _datosDigestivos_AC where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaDigestivos_ES = (from x in _datosDigestivos_ES where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaDigestivos_DI = (from x in _datosDigestivos_DI where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaDigestivos_TI = (from x in _datosDigestivos_TI where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaReproductivos_RE = (from x in _datosReproductivos_RE where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaReproductivos_ME = (from x in _datosReproductivos_ME where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaReproductivos_PIO = (from x in _datosReproductivos_PIO where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaReproductivos_QUI = (from x in _datosReproductivos_QUI where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaReproductivos_CS = (from x in _datosReproductivos_CS where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaRespiratorios_Neu = (from x in _datosRespiratorios_Neu where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaBecerras_Neu = (from x in _datosBecerras_Neu where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaBecerras_Fie = (from x in _datosBecerras_Fie where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaBecerras_Di = (from x in _datosBecerras_Di where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaBecerras_Conj = (from x in _datosBecerras_Conj where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    ValInventario busquedaValInventarios = (from x in _datosValInventarios where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaAbortosVacas = (from x in _datosAbortosVacas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaAbortosVaquillas = (from x in _datosAbortosVaquillas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    #endregion

                    #region Asignar Datos
                    newItem.Ubre_MA = busquedaUbre_MA != null ? busquedaUbre_MA.Vacas : 0;
                    newItem.Ubre_SL = busquedaUbre_SL != null ? busquedaUbre_SL.Vacas : 0;
                    newItem.Metabolicos_FL = busquedaMetabolicos_FL != null ? busquedaMetabolicos_FL.Vacas : 0;
                    newItem.Metabolicos_CET = busquedaMetabolicos_CET != null ? busquedaMetabolicos_CET.Vacas : 0;
                    newItem.Locomotores_BE = busquedaLocomotores_BE != null ? busquedaLocomotores_BE.Vacas : 0;
                    newItem.Locomotores_TRA = busquedaLocomotores_TRA != null ? busquedaLocomotores_TRA.Vacas : 0;
                    newItem.Locomotores_GA = busquedaLocomotores_GA != null ? busquedaLocomotores_GA.Vacas : 0;
                    newItem.Digestivos_AC = busquedaDigestivos_AC != null ? busquedaDigestivos_AC.Vacas : 0;
                    newItem.Digestivos_ES = busquedaDigestivos_ES != null ? busquedaDigestivos_ES.Vacas : 0;
                    newItem.Digestivos_DI = busquedaDigestivos_DI != null ? busquedaDigestivos_DI.Vacas : 0;
                    newItem.Digestivos_TI = busquedaDigestivos_TI != null ? busquedaDigestivos_TI.Vacas : 0;
                    newItem.Reproductivos_RE = busquedaReproductivos_RE != null ? busquedaReproductivos_RE.Vacas : 0;
                    newItem.Reproductivos_ME = busquedaReproductivos_ME != null ? busquedaReproductivos_ME.Vacas : 0;
                    newItem.Reproductivos_PIO = busquedaReproductivos_PIO != null ? busquedaReproductivos_PIO.Vacas : 0;
                    newItem.Reproductivos_QUI = busquedaReproductivos_QUI != null ? busquedaReproductivos_QUI.Vacas : 0;
                    newItem.Reproductivos_CS = busquedaReproductivos_CS != null ? busquedaReproductivos_CS.Vacas : 0;
                    newItem.Respiratorios_Neu = busquedaRespiratorios_Neu != null ? busquedaRespiratorios_Neu.Vacas : 0;
                    newItem.Becerras_Neu = busquedaBecerras_Neu != null ? busquedaBecerras_Neu.Vacas : 0;
                    newItem.Becerras_Fie = busquedaBecerras_Fie != null ? busquedaBecerras_Fie.Vacas : 0;
                    newItem.Becerras_Di = busquedaBecerras_Di != null ? busquedaBecerras_Di.Vacas : 0;
                    newItem.Becerras_Conj = busquedaBecerras_Conj != null ? busquedaBecerras_Conj.Vacas : 0;

                    newItem.Vacas_Pren = busquedaValInventarios != null ? busquedaValInventarios.VacasPreñadas : 0;
                    newItem.Vacas_Vacias = busquedaValInventarios != null ? busquedaValInventarios.VacasVacias : 0;
                    newItem.Vacas_Diag = newItem.Vacas_Pren + newItem.Vacas_Vacias;
                    newItem.Vacas_Porcentaje_Pren = newItem.Vacas_Diag > 0 ? (newItem.Vacas_Pren / newItem.Vacas_Diag) * 100 : 0;
                    newItem.Vacas_Porcentaje_Vacias = newItem.Vacas_Diag > 0 ? (newItem.Vacas_Vacias / newItem.Vacas_Diag) * 100 : 0;

                    newItem.Vaquillas_Pren = busquedaValInventarios != null ? busquedaValInventarios.VaquillasPreñadas : 0;
                    newItem.Vaquillas_Vacias = busquedaValInventarios != null ? busquedaValInventarios.VaquillasVacias : 0;
                    newItem.Vaquillas_Diag = newItem.Vaquillas_Pren + newItem.Vaquillas_Vacias;
                    newItem.Vaquillas_Porcentaje_Pren = newItem.Vaquillas_Diag > 0 ? (newItem.Vaquillas_Pren / newItem.Vaquillas_Diag) * 100 : 0;
                    newItem.Vaquillas_Porcentaje_Vacias = newItem.Vaquillas_Diag > 0 ? (newItem.Vaquillas_Vacias / newItem.Vaquillas_Diag) * 100 : 0;
                    newItem.Abortos_Vaquillas = busquedaAbortosVaquillas != null ? busquedaAbortosVaquillas.Vacas : 0;
                    newItem.Abortos_Vacas = busquedaAbortosVacas != null ? busquedaAbortosVacas.Vacas : 0;
                    #endregion

                    response.Add(newItem);
                }
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "GetReporteHoja4", $"Error in {nameof(GetReporteHoja4)}", ex.Message, ex.Message, ex.StackTrace, guid);
            }

            return response;
        }

        #endregion

        #region  metodos Auxiliares
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

        private DateTime ConvertJulianaToGregoriana(int request)
        {
            return new DateTime(1900, 1, 1).AddDays(request).AddDays(-2);
        }

        /*Carga de Informacion de Hoja 1 y 2*/
        private async Task<bool> CargaInformacionHoja1(gthEntity.CentroTyp centro, Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            var datosproduccionTask = new Task(() => { DatosDeProduccion(fechaInicio, fechaFin); });
            var nivelCribasTask = new Task(() => { NivelCribas(fechaInicio, fechaFin); });
            var inventarioTask = new Task(() => { Inventario(fechaInicio, fechaFin); });
            var calostroTask = new Task(() => { CalostroOrdeño(fechaInicio, fechaFin); });
            var mproducTask = new Task(() => { Mproduc(rancho, fechaInicio, fechaFin); });
            var inventarioAFITask = new Task(() => { InventarioAFIxDIA(fechaInicio, fechaFin); });
            var mproducTask2 = new Task(() => { Mediaproduccion2(fechaInicio, fechaFin); });
            var bacteriologiaTask = new Task(() => { BacteriologiaLeche(fechaInicio, fechaFin); });

            datosproduccionTask.Start();
            nivelCribasTask.Start();
            inventarioTask.Start();
            calostroTask.Start();
            mproducTask.Start();
            inventarioAFITask.Start();
            mproducTask2.Start();
            bacteriologiaTask.Start();

            var info = new List<Task>
            {
                datosproduccionTask,
                nivelCribasTask,
                inventarioTask,
                calostroTask,
                mproducTask,
                inventarioAFITask,
                mproducTask2,
                bacteriologiaTask
            };

            Task.WaitAll(info.ToArray());
            Console.WriteLine("Proceso {0} Ejecutado.", "Carga Informacion Hoja 1 y 2");

            return true;
        }

        private async Task<bool> CargaInformacionHoja3(gthEntity.CentroTyp centro, Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            var jaulasTask = new Task(() => { DatosJaulasHoja3(fechaInicio, fechaFin); });
            var desteteTask = new Task(() => { DatosDesteHoja3(fechaInicio, fechaFin); });
            var vaquillasTask = new Task(() => { DatosVaquillasHoja3(fechaInicio, fechaFin); });
            var vacasTask = new Task(() => { DatosVacasHoja3(fechaInicio, fechaFin); });
            var partosVaqTask = new Task(() => { DatosPartosVaquillasHoja3(fechaInicio, fechaFin); });
            var partosVacasTask = new Task(() => { DatosPartosVacasHoja3(fechaInicio, fechaFin); });
            var vacasMuertasTask = new Task(() => { DatosVacasMuertasHoja3(fechaInicio, fechaFin); });

            jaulasTask.Start();
            desteteTask.Start();
            vaquillasTask.Start();
            vacasTask.Start();
            partosVaqTask.Start();
            partosVacasTask.Start();
            vacasMuertasTask.Start();

            var info = new List<Task>
            {
                jaulasTask,
                desteteTask,
                vaquillasTask,
                vacasTask,
                partosVaqTask,
                partosVacasTask,
                vacasMuertasTask
            };

            Task.WaitAll(info.ToArray());
            Console.WriteLine("Proceso {0} Ejecutado.", "Carga Informacion Hoja 3");

            return true;
        }

        private async Task<bool> CargaInformacionHoja4(gthEntity.CentroTyp centro, Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            var ubresTask = new Task(() => { DatosUbres(fechaInicio, fechaFin); });
            var metabolicosTask = new Task(() => { DatosMetabolicos(fechaInicio, fechaFin); });
            var locomotoresTask = new Task(() => { DatosLocomotores(fechaInicio, fechaFin); });
            var digestivosTask = new Task(() => { DatosDigestivos(fechaInicio, fechaFin); });
            var reproductivosTask = new Task(() => { DatosReproductivos(fechaInicio, fechaFin); });
            var enfermedadesTask = new Task(() => { DatosEnFermedadesBecerras(fechaInicio, fechaFin); });
            var valInventarioTask = new Task(() => { DatosValInventarios(fechaInicio, fechaFin); });
            var abortosTask = new Task(() => { DatosAbortos(fechaInicio, fechaFin); });            

            ubresTask.Start();
            metabolicosTask.Start();
            locomotoresTask.Start();
            digestivosTask.Start();
            reproductivosTask.Start();
            enfermedadesTask.Start();
            valInventarioTask.Start();
            abortosTask.Start();

            var info = new List<Task>
            {
                ubresTask,
                metabolicosTask,
                locomotoresTask,
                digestivosTask,
                reproductivosTask,
                enfermedadesTask,
                valInventarioTask,
                abortosTask
            };

            Task.WaitAll(info.ToArray());
            Console.WriteLine("Proceso {0} Ejecutado.", "Carga Informacion Hoja 4");

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

        private async Task<bool> PromediosHoja3(DateTime fechaInicio, DateTime fechaFin)
        {

            Console.WriteLine("Proceso {0} Ejecutado.", "PromediosHoja3");
            return true;
        }

        private async Task<bool> PromediosHoja4(DateTime fechaInicio, DateTime fechaFin)
        {

            Console.WriteLine("Proceso {0} Ejecutado.", "PromediosHoja4");
            return true;
        }

        #endregion

        #region Consultas Asincronas

        #region datos teoricos
        private async Task<bool> IndicadoresTeoricoProduccion(List<gthEntity.CentroTyp> centro, DateTime fechaInicio, DateTime fechaFin)
        {
            indicadorProduccion = consumoModel.GetIndicadorReporteTeorico(centro, etapaProduccion, fechaInicio, fechaFin);
            Console.WriteLine("Proceso {0} Ejecutado.", "IndicadoresTeoricoSecas");
            return true;
        }

        private async Task<bool> IndicadoresTeoricoSecas(List<gthEntity.CentroTyp> centro, DateTime fechaInicio, DateTime fechaFin)
        {
            indicadorSecas = consumoModel.GetIndicadorReporteTeorico(centro, etapaSecas, fechaInicio, fechaFin);
            Console.WriteLine("Proceso {0} Ejecutado.", "IndicadoresTeoricoSecas");
            return true;
        }

        private async Task<bool> IndicadoresTeoricoReto(List<gthEntity.CentroTyp> centro, DateTime fechaInicio, DateTime fechaFin)
        {
            indicadorReto = consumoModel.GetIndicadorReporteTeorico(centro, etapaReto, fechaInicio, fechaFin);
            Console.WriteLine("Proceso {0} Ejecutado.", "IndicadoresTeoricoReto");
            return true;
        }

        private async Task<bool> IndicadoresTeoricoCrecimiento(List<gthEntity.CentroTyp> centro, DateTime fechaInicio, DateTime fechaFin)
        {
            indicadorCrecimiento = consumoModel.GetIndicadorReporteTeorico(centro, etapaCrecimiento, fechaInicio, fechaFin);
            Console.WriteLine("Proceso {0} Ejecutado.", "IndicadoresTeoricoCrecimiento");
            return true;
        }

        private async Task<bool> IndicadoresTeoricoDesarrollo(List<gthEntity.CentroTyp> centro, DateTime fechaInicio, DateTime fechaFin)
        {
            indicadorDesarrollo = consumoModel.GetIndicadorReporteTeorico(centro, etapaDesarrollo, fechaInicio, fechaFin);
           Console.WriteLine("Proceso {0} Ejecutado.", "IndicadoresTeoricoDesarrollo");
            return true;
        }

        private async Task<bool> IndicadoresTeoricoVaquillas(List<gthEntity.CentroTyp> centro, DateTime fechaInicio, DateTime fechaFin)
        {
            indicadorVaquillas = consumoModel.GetIndicadorReporteTeorico(centro, etapaVaquillas, fechaInicio, fechaFin);
            Console.WriteLine("Proceso {0} Ejecutado.", "IndicadoresTeoricoVaquillas");
            return true;
        }
        #endregion

        #region datos alimentacion
        private async Task<bool> IndicadorProduccion(List<gthEntity.CentroTyp> centro, DateTime fechaInicio, DateTime fechaFin)
        {
            try
            {
                promedioProduccion = consumoModel.GetIndicadorReportePeriodo(centro, etapaProduccion, fechaInicio, fechaFin);
                promedioAnteriorProduccion = consumoModel.GetIndicadorReportePeriodo(centro, etapaProduccion, fechaInicio.AddYears(-1), fechaFin.AddYears(-1));
                Console.WriteLine("Proceso {0} Ejecutado.", "IndicadorProduccion");
            }
            catch { }
            return true;
        }

        private async Task<bool> IndicadorSecas(List<gthEntity.CentroTyp> centro, DateTime fechaInicio, DateTime fechaFin)
        {
            try
            {
                promedioSecas = consumoModel.GetIndicadorReportePeriodo(centro, etapaSecas, fechaInicio, fechaFin);
                promedioAnteriorSecas = consumoModel.GetIndicadorReportePeriodo(centro, etapaSecas, fechaInicio.AddYears(-1), fechaFin.AddYears(-1));
                Console.WriteLine("Proceso {0} Ejecutado.", "IndicadorSecas");
            }
            catch { }
            return true;
        }

        private async Task<bool> IndicadorReto(List<gthEntity.CentroTyp> centro, DateTime fechaInicio, DateTime fechaFin)
        {
            try
            {
                promedioReto = consumoModel.GetIndicadorReportePeriodo(centro, etapaReto, fechaInicio, fechaFin);
                promedioAnteriorReto = consumoModel.GetIndicadorReportePeriodo(centro, etapaReto, fechaInicio.AddYears(-1), fechaFin.AddYears(-1));
                Console.WriteLine("Proceso {0} Ejecutado.", "IndicadorReto");
            }
            catch { }
            return true;
        }

        private async Task<bool> IndicadorJaulas(List<gthEntity.CentroTyp> centro, DateTime fechaInicio, DateTime fechaFin)
        {
            try
            {
                promedioJaulas = consumoModel.GetIndicadorReportePeriodo(centro, etapaJaulas, fechaInicio, fechaFin);
                promedioAnteriorJaulas = consumoModel.GetIndicadorReportePeriodo(centro, etapaJaulas, fechaInicio.AddYears(-1), fechaFin.AddYears(-1));
                Console.WriteLine("Proceso {0} Ejecutado.", "IndicadorJaulas");
            }
            catch { }
            return true;
        }

        private async Task<bool> IndicadorCrecimiento(List<gthEntity.CentroTyp> centro, DateTime fechaInicio, DateTime fechaFin)
        {
            try
            {
                promedioCrecimiento = consumoModel.GetIndicadorReportePeriodo(centro, etapaCrecimiento, fechaInicio, fechaFin);
                promedioAnteriorCrecimiento = consumoModel.GetIndicadorReportePeriodo(centro, etapaCrecimiento, fechaInicio.AddYears(-1), fechaFin.AddYears(-1));
                Console.WriteLine("Proceso {0} Ejecutado.", "IndicadorCrecimiento");
            }
            catch { }
            return true;
        }

        private async Task<bool> IndicadorDesarrollo(List<gthEntity.CentroTyp> centro, DateTime fechaInicio, DateTime fechaFin)
        {
            try
            {
                promedioDesarrollo = consumoModel.GetIndicadorReportePeriodo(centro, etapaDesarrollo, fechaInicio, fechaFin);
                promedioAnteriorDesarrollo = consumoModel.GetIndicadorReportePeriodo(centro, etapaDesarrollo, fechaInicio.AddYears(-1), fechaFin.AddYears(-1));
                Console.WriteLine("Proceso {0} Ejecutado.", "IndicadorDesarrollo");
            }
            catch { }
            return true;
        }

        private async Task<bool> IndicadorVaquillas(List<gthEntity.CentroTyp> centro, DateTime fechaInicio, DateTime fechaFin)
        {
            try
            {
                promedioVaquillas = consumoModel.GetIndicadorReportePeriodo(centro, etapaVaquillas, fechaInicio, fechaFin);
                promedioAnteriorVaquillas = consumoModel.GetIndicadorReportePeriodo(centro, etapaVaquillas, fechaInicio.AddYears(-1), fechaFin.AddYears(-1));
                Console.WriteLine("Proceso {0} Ejecutado.", "IndicadorVaquillas");
            }
            catch { }
            return true;
        }
        #endregion

        #region datos Hoja 1 y 2 
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

        private async Task<bool> BacteriologiaLeche(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosBacteriologia = GetBacteriologiaLeche(fechaInicio, fechaFin);
            Console.WriteLine("Proceso {0} Ejecutado.", "BacteriologiaLeche");
            return true;
        }

        private async Task<bool> PrecioLecheTask(gthEntity.CentroTyp centro, DateTime fechaInicio, DateTime fechaFin)
        {
            _precioLeche = GetPrecioLeche(centro.CentroClave, fechaInicio, fechaFin);
            _precioLecheAñoAnt = GetPrecioLeche(centro.CentroClave, fechaInicio.AddYears(-1), fechaFin.AddYears(-1));
            Console.WriteLine("Proceso {0} Ejecutado.", "PrecioLeche");
            return true;
        }

        private async Task<bool> LecheProducidaTask(gthEntity.CentroTyp centro, DateTime fechaInicio, DateTime fechaFin)
        {
            _lecheProducida = GetLecheProducida(fechaInicio, fechaFin);
            _lecheProducidaAñoAnt = GetLecheProducida(fechaInicio.AddYears(-1), fechaFin.AddYears(-1));
            return true;
        }
        #endregion

        #region  datos hoja 3
        private async Task<bool> DatosJaulasHoja3(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosJaulasFuera = GetDatosDesecho(fechaInicio, fechaFin, @"AND LACTANCIA = 0 AND EDOVAC = 15");
            _datosJaulasVivas = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo <> 1 AND destino <> 2 AND (edovac = 1 OR edovac = 10) ");
            _datosJaulasMuertas = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 1 AND (edovac = 1 OR edovac = 10) ");
            return true;
        }

        private async Task<bool> DatosDesteHoja3(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosDesteteVivas = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo <> 1 AND destino <> 2 AND (edovac = 2 OR edovac = 7) ");
            _datosDesteteMuertas = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 1 AND (edovac = 2 OR edovac = 7) ");
            return true;
        }

        private async Task<bool> DatosVaquillasHoja3(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosVaquillasMuertas = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 1 AND destino = 1  AND (edovac = 3 OR edovac = 8 OR edovac = 9) ");
            _datosVaquillasUrgencia = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 2 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ");
            _datosVaquillasDelgadas = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 3 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ");
            _datosVaquillasRegulares = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 4 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ");
            _datosVaquillasGordas = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 5 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ");
            _datosVaquillasOtros = GetDatosDesecho(fechaInicio, fechaFin, @"  AND vacvaq = 2  AND motivo > 5 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ");
            return true;
        }

        private async Task<bool> DatosVacasHoja3(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosVacasMuertas = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo = 1 AND destino = 1 AND (edovac > 3) ");
            _datosVacasUrgencia = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo = 2 AND destino = 1 AND (edovac > 3) ");
            _datosVacasDelgadas = GetDatosDesecho(fechaInicio, fechaFin, @"  AND vacvaq = 1  AND motivo = 3 AND destino = 1 AND (edovac > 3) ");
            _datosVacasRegulares = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo = 4 AND destino = 1 AND (edovac > 3) ");
            _datosVacasGordas = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo = 5 AND destino = 1 AND (edovac > 3) ");
            _datosVacasOtros = GetDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo > 5 AND destino = 1 AND (edovac > 3) ");
            return true;
        }

        private async Task<bool> DatosPartosVaquillasHoja3(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosPartosVaquillasND = GetDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2  AND (tipopar = 1 or tipopar = 3) ");
            _datosPartosVaquillasAbortos = GetDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2  AND (tipopar = 2) ");
            _datosPartosVaquillasHembra = GetDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2  AND criaviva = 1  AND criasexo = 1 ");
            _datosPartosVaquillasMacho = GetDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2  AND criaviva = 1  AND criasexo = 2 ");
            return true;
        }

        private async Task<bool> DatosPartosVacasHoja3(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosPartosVacasND = GetDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 1  AND (tipopar = 1 or tipopar = 3) ");
            _datosPartosVacasAbortos = GetDatosNacimiento(fechaInicio, fechaFin, @"  AND vacvaq = 1  AND (tipopar = 2) ");
            _datosPartosVacasHembra = GetDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 1  AND criaviva = 1  AND criasexo = 1 ");
            _datosPartosVacasMacho = GetDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 1  AND criaviva = 1  AND criasexo = 2 ");

            _datosPartosSinCria = GetDatosNacimiento(fechaInicio, fechaFin, @" AND tipopar <> 2 AND criaviva = 2  AND criasexo = 3 ");
            _datosPartosVacas = GetDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 1 ");
            _datosPartosVaquillas = GetDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2 ");
            return true;
        }

        private async Task<bool> DatosVacasMuertasHoja3(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosMuertasDia = GetDatosNacimientoMuertas(fechaInicio, fechaFin, @" AND criaviva = 2 AND (tipopar = 1 or tipopar = 3) AND criasexo < 3 AND dianoche = 1 ");
            _datosMuertasNoche = GetDatosNacimientoMuertas(fechaInicio, fechaFin, @" AND criaviva = 2 AND (tipopar = 1 or tipopar = 3) AND criasexo < 3 AND dianoche = 2 ");
            return true;
        }
        #endregion

        #region  datos hoja 4
        private async Task<bool> DatosUbres(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosUbre_MA = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 1 AND enfdetalle = 1 ");
            _datosUbre_SL = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 1 AND enfdetalle = 2 ");
            return true;
        }

        private async Task<bool> DatosMetabolicos(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosMetabolicos_FL = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 5 AND enfdetalle = 1 ");
            _datosMetabolicos_CET = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 5 AND enfdetalle = 2 ");
            return true;
        }

        private async Task<bool> DatosLocomotores(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosLocomotores_BE = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 4 AND enfdetalle = 1 ");
            _datosLocomotores_TRA = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 4 AND enfdetalle = 2 ");
            _datosLocomotores_GA = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 4 AND enfdetalle = 3 ");
            return true;
        }

        private async Task<bool> DatosDigestivos(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosDigestivos_AC = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 1 ");
            _datosDigestivos_ES = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 2 ");
            _datosDigestivos_DI = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 3 ");
            _datosDigestivos_TI = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 4 ");
            return true;
        }

        private async Task<bool> DatosReproductivos(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosReproductivos_RE = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 1 ");
            _datosReproductivos_ME = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 2 ");
            _datosReproductivos_PIO = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 3 ");
            _datosReproductivos_QUI = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 4 ");
            _datosReproductivos_CS = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 5 ");
            return true;
        }

        private async Task<bool> DatosEnFermedadesBecerras(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosRespiratorios_Neu = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 7 AND enfdetalle = 1 ");
            _datosBecerras_Neu = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 2  AND enfermedad = 8 AND enfdetalle = 1 ");
            _datosBecerras_Fie = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 2  AND enfermedad = 8 AND enfdetalle = 2 ");
            _datosBecerras_Di = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 2  AND enfermedad = 8 AND enfdetalle = 3 ");
            _datosBecerras_Conj = GetDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 2  AND enfermedad = 8 AND enfdetalle = 4 ");
            return true;
        }

        private async Task<bool> DatosValInventarios(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosValInventarios = GetDatosValInventario(fechaInicio, fechaFin);
            return true;
        }

        private async Task<bool> DatosAbortos(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosAbortosVacas = GetDatosAbortos(fechaInicio, fechaFin, 1);
            _datosAbortosVaquillas = GetDatosAbortos(fechaInicio, fechaFin, 2);
            return true;
        }
        #endregion

        #endregion


        #region metodos consulta x dia

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

        private List<BacteriologiaLeche> GetBacteriologiaLeche(DateTime fechaInicio, DateTime fechaFin)
        {
            List<BacteriologiaLeche> response = new List<BacteriologiaLeche>();
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

            try
            {
                string query = @"SELECT  Inidicadores.FECHA                                     AS FECHA       
										,Inidicadores.PRECIO_BACTERIOLOGIA
								FROM
								(
									SELECT  CDATE(Inc_FECHA)         AS FECHA
											,Inc_GRASA                AS GRASA
											,Inc_PROTEINA             AS PROTEINA
											,Inc_CELULAS              AS CELULASSOMATICAS
											,Inc_BACTERIOLOGIA        AS BACTERIOLOGIA
											,Inc_TEMPERATURA          AS TEMPERATURA
											,Inc_ANTIBIOTICO          AS ANTIBIOTICO
											,Inc_SEDIMIENTO           AS SEDIMIENTOS
											,Inc_CRIOSCOPIA           AS CRIOSCOPIA
											,Inc_BRUCELA              AS BRUCELA
											,Inc_FEDERAL              AS FEDERAL
											,Inc_PRECIO_BACTERIOLOGIA AS PRECIO_BACTERIOLOGIA
									FROM INCENTIVOSLECHE
									WHERE Inc_FECHA BETWEEN #@fechaIni# AND #@fechaFin# 
								) Inidicadores
								LEFT JOIN
								(
									SELECT  CDATE(FECHA)       AS FECHAS
											,SUM(LITROSXTANQUE) AS LITROSPORDIA
									FROM DPRODUC
									WHERE FECHA BETWEEN #@fechaIni# AND #@fechaFin#
									GROUP BY  FECHA
								)LitrosLeche
								ON Inidicadores.FECHA = LitrosLeche.FECHAS
								ORDER BY Inidicadores.FECHA";
                query = query.Replace("@fechaIni", fechaInicio.ToString("yyyy/MM/dd")).Replace("@fechaFin", fechaFin.ToString("yyyy/MM/dd"));

                db.Conectar();
                db.CrearComando(query, tipoComandoOleDb.query);
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new BacteriologiaLeche()
                {
                    Fecha = x["FECHA"] != DBNull.Value ? Convert.ToDateTime(x["FECHA"]) : new DateTime(1, 1, 1),
                    Bacteriologia = x["PRECIO_BACTERIOLOGIA"] != DBNull.Value ? Convert.ToDecimal(x["PRECIO_BACTERIOLOGIA"]) : 0
                }).ToList();
            }
            catch (DbException ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "BacteriologiaLeche", $"Error in {nameof(GetBacteriologiaLeche)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
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

        private decimal GetPrecioLeche(int centroId, DateTime fechaInicio, DateTime fechaFin)
        {
            decimal response = 0;

            if (EsPrecioLecheFacturado(fechaFin))
                response = GetPrecioLecheFacturado(fechaInicio, fechaFin);
            else
                response = GetPrecioLecheCalculado(centroId, fechaFin, fechaInicio);


            return response;
        }

        private decimal GetLecheProducida(DateTime fechaInicio, DateTime fechaFin)
        {
            decimal response = 0;
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
                    response = dt.Rows[0]["LecheProd"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["LecheProd"]) : 0;
            }
            catch (DbException ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "PrecioLecheFacturado", $"Error in {nameof(GetLecheProducida)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        #endregion

        #region hoja 3 y 4
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
                InsLog("ReporteDiario", "GetDatosDesecho", $"Error in {nameof(GetDatosDesecho)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
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
                InsLog("ReporteDiario", "GetDatosNacimiento", $"Error in {nameof(GetDatosNacimiento)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
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
                InsLog("ReporteDiario", "GetDatosNacimientoMuertas", $"Error in {nameof(GetDatosNacimientoMuertas)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
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
                InsLog("ReporteDiario", "GetDatosAbortos", $"Error in {nameof(GetDatosAbortos)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

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
            catch (DbException ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "GetDatosSalud", $"Error in {nameof(GetDatosSalud)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
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
                InsLog("ReporteDiario", "GetDatosValInventario", $"Error in {nameof(GetDatosValInventario)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        #endregion

        #endregion

        #region metodos para promedios      
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
        public Hoja3 GetTotalReporteHoja3(DateTime fechaInicio, DateTime fechaFin, CalostroYOrdeña totalCalostro, string dia)
        {
            Hoja3 response = new Hoja3() { Dia = dia };

            try
            {
                #region Obtener Datos                
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
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "GetTotalReporteHoja3", $"Error in {nameof(GetTotalReporteHoja3)}", ex.Message, ex.Message, ex.StackTrace, guid);
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
            catch (DbException ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "GetTotalDatosDesecho", $"Error in {nameof(GetTotalDatosDesecho)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
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
            catch (DbException ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "GetTotalDatosNacimiento", $"Error in {nameof(GetTotalDatosNacimiento)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
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
            catch (DbException ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "GetTotalDatosNacimientoMuertas", $"Error in {nameof(GetTotalDatosNacimientoMuertas)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
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
            catch (DbException ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "GetTotalDatosAbortos", $"Error in {nameof(GetTotalDatosAbortos)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        #endregion

        #region Hoja 4
        public Hoja4 GetTotalHoja4(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, string dia)
        {
            Hoja4 response = new Hoja4();
            response.Dia = dia;

            try
            {
                #region obtener datos
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
                #endregion

                #region asignar datos
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
                #endregion
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "GetTotalHoja4", $"Error in {nameof(GetTotalHoja4)}.", ex.Message, ex.Message, ex.StackTrace, guid);
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
            catch (DbException ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "GetTotalDatosSalud", $"Error in {nameof(GetTotalDatosSalud)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
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
            catch (DbException ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "GetTotalDatosValInventario", $"Error in {nameof(GetTotalDatosValInventario)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }
        #endregion

        private async Task<bool> PromedioMproduc(Rancho rancho, DateTime fechaInicio, DateTime fechaFin)
        {
            _promedioHoja1 = rancho.NoIdReal ? GetPromedioHoja1NoIdReal(fechaInicio, fechaFin) : GetPromedioHoja1NoId(fechaInicio, fechaFin);
            _promedioHoja1AñoAnt = rancho.NoIdReal ? GetPromedioHoja1NoIdReal(fechaInicio.AddYears(-1), fechaFin.AddYears(-1)) : GetPromedioHoja1NoId(fechaInicio.AddYears(-1), fechaFin.AddYears(-1));
            Console.WriteLine("Proceso {0} Ejecutado.", "PromedioMproduc");

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

        private async Task<bool> TotalInventarioAFI(DateTime fechaInicio, DateTime fechaFin)
        {
            _promedioInventarioAFI = GetPromedioInventariosAFI(fechaInicio, fechaFin);
            _promedioInventarioAFIAñoAnt = GetPromedioInventariosAFI(fechaInicio.AddYears(-1), fechaFin.AddYears(-1));
            Console.WriteLine("Proceso {0} Ejecutado.", "TotalInventarioAFI");
            return true;
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

        private InventarioAfiXDia GetPromedioInventariosAFI(DateTime fechaInicio, DateTime fechaFin)
        {
            InventarioAfiXDia response = new InventarioAfiXDia();
            ModeloDatosOleDb db = new ModeloDatosOleDb(_oleDbConnection);

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


        #endregion

        #region diferencias
        public Hoja1 DiferenciaHoja1(Hoja1 promedio, Hoja1 añoAnterior)
        {
            Hoja1 diferencia = new Hoja1() { Dia = "DIF #" };

            try
            {
                diferencia.Ordeño = promedio.Ordeño - añoAnterior.Ordeño;
                diferencia.Secas = promedio.Secas - añoAnterior.Secas;
                diferencia.Hato = promedio.Hato - añoAnterior.Hato;
                diferencia.Porcentaje_Lact = promedio.Porcentaje_Lact - añoAnterior.Porcentaje_Lact;
                diferencia.Porcentaje_Prot = promedio.Porcentaje_Prot - añoAnterior.Porcentaje_Prot;
                diferencia.Urea = promedio.Urea - añoAnterior.Urea;
                diferencia.Porcentaje_Grasa = promedio.Porcentaje_Grasa - añoAnterior.Porcentaje_Grasa;
                diferencia.CCS = promedio.CCS - añoAnterior.CCS;
                diferencia.CTD = promedio.CTD - añoAnterior.CTD;
                diferencia.DEC = null;

                diferencia.Leche = promedio.Leche - añoAnterior.Leche;
                diferencia.Antib = promedio.Antib - añoAnterior.Antib;
                diferencia.Media = promedio.Media - añoAnterior.Media;
                diferencia.Total = promedio.Total - añoAnterior.Total;
                diferencia.DEL = promedio.DEL - añoAnterior.DEL;
                diferencia.ANT = promedio.ANT - añoAnterior.ANT;

                diferencia.EA = promedio.EA - añoAnterior.EA;
                diferencia.ILCA = promedio.ILCA - añoAnterior.ILCA;
                diferencia.IC = promedio.IC - añoAnterior.IC;
                diferencia.Costo_Litro = promedio.Costo_Litro - añoAnterior.Costo_Litro;
                diferencia.MH = promedio.MH - añoAnterior.MH;
                diferencia.Porcentaje_MS = promedio.Porcentaje_MS - añoAnterior.Porcentaje_MS;
                diferencia.MS = promedio.MS - añoAnterior.MS;
                diferencia.SA = promedio.SA - añoAnterior.SA;
                diferencia.MSS = promedio.MSS - añoAnterior.MSS;
                diferencia.EAS = promedio.EAS - añoAnterior.EAS;
                diferencia.Porcentaje_Sob = promedio.Porcentaje_Sob - añoAnterior.Porcentaje_Sob;
                diferencia.Costo_Prod = promedio.Costo_Prod - añoAnterior.Costo_Prod;
                diferencia.Costo_MS = promedio.Costo_MS - añoAnterior.Costo_MS;

                diferencia.Cribas_N1 = promedio.Cribas_N1 - añoAnterior.Cribas_N1;
                diferencia.Cribas_N2 = promedio.Cribas_N2 - añoAnterior.Cribas_N2;
                diferencia.Cribas_N3 = promedio.Cribas_N3 - añoAnterior.Cribas_N3;
                diferencia.Cribas_N4 = promedio.Cribas_N4 - añoAnterior.Cribas_N4;

                diferencia.NoID_Ses1 = promedio.NoID_Ses1 - añoAnterior.NoID_Ses1;
                diferencia.NoID_Ses2 = promedio.NoID_Ses2 - añoAnterior.NoID_Ses2;
                diferencia.NoID_Ses3 = promedio.NoID_Ses3 - añoAnterior.NoID_Ses3;

                diferencia.SES1 = promedio.SES1 - añoAnterior.SES1;
                diferencia.SES2 = promedio.SES2 - añoAnterior.SES2;
                diferencia.SES3 = promedio.SES3 - añoAnterior.SES3;

                diferencia.Porcentaje_Revueltas = promedio.Porcentaje_Revueltas - añoAnterior.Porcentaje_Revueltas;
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "DiferenciaHoja1", $"Error in {nameof(DiferenciaHoja1)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }

            return diferencia;
        }

        public Hoja2 DiferenciaHoja2(Hoja2 promedio, Hoja2 promedioAñoAnt)
        {
            Hoja2 difencia = new Hoja2() { Dia = "DIF #" };

            try
            {
                #region Crianza
                difencia.Crianza_Inventario = promedio.Crianza_Inventario != null && promedioAñoAnt.Crianza_Inventario != null ? promedio.Crianza_Inventario - promedioAñoAnt.Crianza_Inventario : 0;
                #endregion

                #region Jaulas 
                difencia.Jaulas_Inventario = promedio.Jaulas_Inventario != null && promedioAñoAnt.Jaulas_Inventario != null ? promedio.Jaulas_Inventario - promedioAñoAnt.Jaulas_Inventario : 0;
                difencia.Jaulas_Costo = promedio.Jaulas_Costo != null && promedioAñoAnt.Jaulas_Costo != null ? promedio.Jaulas_Costo - promedioAñoAnt.Jaulas_Costo : 0;
                #endregion

                #region Crecimiento 
                difencia.Crecimiento_Inventario = promedio.Crecimiento_Inventario != null && promedioAñoAnt.Crecimiento_Inventario != null ? promedio.Crecimiento_Inventario - promedioAñoAnt.Crecimiento_Inventario : 0;
                difencia.Crecimiento_MH = promedio.Crecimiento_MH != null && promedioAñoAnt.Crecimiento_MH != null ? promedio.Crecimiento_MH - promedioAñoAnt.Crecimiento_MH : 0;
                difencia.Crecimiento_Costo = promedio.Crecimiento_Costo != null && promedioAñoAnt.Crecimiento_Costo != null ? promedio.Crecimiento_Costo - promedioAñoAnt.Crecimiento_Costo : 0;
                difencia.Crecimiento_PorcentajeMS = promedio.Crecimiento_PorcentajeMS != null && promedioAñoAnt.Crecimiento_PorcentajeMS != null ? promedio.Crecimiento_PorcentajeMS - promedioAñoAnt.Crecimiento_PorcentajeMS : 0;
                difencia.Crecimiento_MS = promedio.Crecimiento_MS != null && promedioAñoAnt.Crecimiento_MS != null ? promedio.Crecimiento_MS - promedioAñoAnt.Crecimiento_MS : 0;
                difencia.Crecimiento_CostoMS = promedio.Crecimiento_CostoMS != null && promedioAñoAnt.Crecimiento_CostoMS != null ? promedio.Crecimiento_CostoMS - promedioAñoAnt.Crecimiento_CostoMS : 0;
                #endregion

                #region Desarrollo 
                difencia.Desarrollo_Inventario = promedio.Desarrollo_Inventario != null && promedioAñoAnt.Desarrollo_Inventario != null ? promedio.Desarrollo_Inventario - promedioAñoAnt.Desarrollo_Inventario : 0;
                difencia.Desarrollo_MH = promedio.Desarrollo_MH != null && promedioAñoAnt.Desarrollo_MH != null ? promedio.Desarrollo_MH - promedioAñoAnt.Desarrollo_MH : 0;
                difencia.Desarrollo_Costo = promedio.Desarrollo_Costo != null && promedioAñoAnt.Desarrollo_Costo != null ? promedio.Desarrollo_Costo - promedioAñoAnt.Desarrollo_Costo : 0;
                difencia.Desarrollo_PorcentajeMS = promedio.Desarrollo_PorcentajeMS != null && promedioAñoAnt.Desarrollo_PorcentajeMS != null ? promedio.Desarrollo_PorcentajeMS - promedioAñoAnt.Desarrollo_PorcentajeMS : 0;
                difencia.Desarrollo_MS = promedio.Desarrollo_MS != null && promedioAñoAnt.Desarrollo_MS != null ? promedio.Desarrollo_MS - promedioAñoAnt.Desarrollo_MS : 0;
                difencia.Desarrollo_CostoMS = promedio.Desarrollo_CostoMS != null && promedioAñoAnt.Desarrollo_CostoMS != null ? promedio.Desarrollo_CostoMS - promedioAñoAnt.Desarrollo_CostoMS : 0;
                #endregion

                #region Vaquillas 
                difencia.Vaquillas_Inventario = promedio.Vaquillas_Inventario != null && promedioAñoAnt.Vaquillas_Inventario != null ? promedio.Vaquillas_Inventario - promedioAñoAnt.Vaquillas_Inventario : 0;
                difencia.Vaquillas_MH = promedio.Vaquillas_MH != null && promedioAñoAnt.Vaquillas_MH != null ? promedio.Vaquillas_MH - promedioAñoAnt.Vaquillas_MH : 0;
                difencia.Vaquillas_Costo = promedio.Vaquillas_Costo != null && promedioAñoAnt.Vaquillas_Costo != null ? promedio.Vaquillas_Costo - promedioAñoAnt.Vaquillas_Costo : 0;
                difencia.Vaquillas_PorcentajeMS = promedio.Vaquillas_PorcentajeMS != null && promedioAñoAnt.Vaquillas_PorcentajeMS != null ? promedio.Vaquillas_PorcentajeMS - promedioAñoAnt.Vaquillas_PorcentajeMS : 0;
                difencia.Vaquillas_MS = promedio.Vaquillas_MS != null && promedioAñoAnt.Vaquillas_MS != null ? promedio.Vaquillas_MS - promedioAñoAnt.Vaquillas_MS : 0;
                difencia.Vaquillas_CostoMS = promedio.Vaquillas_CostoMS != null && promedioAñoAnt.Vaquillas_CostoMS != null ? promedio.Vaquillas_CostoMS - promedioAñoAnt.Vaquillas_CostoMS : 0;
                #endregion

                #region Secas 
                difencia.Secas_Inventario = promedio.Secas_Inventario != null && promedioAñoAnt.Secas_Inventario != null ? promedio.Secas_Inventario - promedioAñoAnt.Secas_Inventario : 0;
                difencia.Secas_MH = promedio.Secas_MH != null && promedioAñoAnt.Secas_MH != null ? promedio.Secas_MH - promedioAñoAnt.Secas_MH : 0;
                difencia.Secas_PorcentajeMS = promedio.Secas_PorcentajeMS != null && promedioAñoAnt.Secas_PorcentajeMS != null ? promedio.Secas_PorcentajeMS - promedioAñoAnt.Secas_PorcentajeMS : 0;
                difencia.Secas_MS = promedio.Secas_MS != null && promedioAñoAnt.Secas_MS != null ? promedio.Secas_MS - promedioAñoAnt.Secas_MS : 0;
                difencia.Secas_SA = promedio.Secas_SA != null && promedioAñoAnt.Secas_SA != null ? promedio.Secas_SA - promedioAñoAnt.Secas_SA : 0;
                difencia.Secas_Mss = promedio.Secas_Mss != null && promedioAñoAnt.Secas_Mss != null ? promedio.Secas_Mss - promedioAñoAnt.Secas_Mss : 0;
                difencia.Secas_PorcentajeSob = promedio.Secas_PorcentajeSob != null && promedioAñoAnt.Secas_PorcentajeSob != null ? promedio.Secas_PorcentajeSob - promedioAñoAnt.Secas_PorcentajeSob : 0;
                difencia.Secas_Costo = promedio.Secas_Costo != null && promedioAñoAnt.Secas_Costo != null ? promedio.Secas_Costo - promedioAñoAnt.Secas_Costo : 0;
                difencia.Secas_CostoMS = promedio.Secas_CostoMS != null && promedioAñoAnt.Secas_CostoMS != null ? promedio.Secas_CostoMS - promedioAñoAnt.Secas_CostoMS : 0;
                #endregion

                #region Reto 
                difencia.Reto_Inventario = promedio.Reto_Inventario != null && promedioAñoAnt.Reto_Inventario != null ? promedio.Reto_Inventario - promedioAñoAnt.Reto_Inventario : 0;
                difencia.Reto_MH = promedio.Reto_MH != null && promedioAñoAnt.Reto_MH != null ? promedio.Reto_MH - promedioAñoAnt.Reto_MH : 0;
                difencia.Reto_PorcentajeMS = promedio.Reto_PorcentajeMS != null && promedioAñoAnt.Reto_PorcentajeMS != null ? promedio.Reto_PorcentajeMS - promedioAñoAnt.Reto_PorcentajeMS : 0;
                difencia.Reto_MS = promedio.Reto_MS != null && promedioAñoAnt.Reto_MS != null ? promedio.Reto_MS - promedioAñoAnt.Reto_MS : 0;
                difencia.Reto_SA = promedio.Reto_SA != null && promedioAñoAnt.Reto_SA != null ? promedio.Reto_SA - promedioAñoAnt.Reto_SA : 0;
                difencia.Reto_Mss = promedio.Reto_Mss != null && promedioAñoAnt.Reto_Mss != null ? promedio.Reto_Mss - promedioAñoAnt.Reto_Mss : 0;
                difencia.Reto_PorcentajeSob = promedio.Reto_PorcentajeSob != null && promedioAñoAnt.Reto_PorcentajeSob != null ? promedio.Reto_PorcentajeSob - promedioAñoAnt.Reto_PorcentajeSob : 0;
                difencia.Reto_Costo = promedio.Reto_Costo != null && promedioAñoAnt.Reto_Costo != null ? promedio.Reto_Costo - promedioAñoAnt.Reto_Costo : 0;
                difencia.Reto_CostoMS = promedio.Reto_CostoMS != null && promedioAñoAnt.Reto_CostoMS != null ? promedio.Reto_CostoMS - promedioAñoAnt.Reto_CostoMS : 0;
                #endregion

                #region utilidad
                difencia.Inventario_Total = promedio.Inventario_Total != null && promedioAñoAnt.Inventario_Total != null ? promedio.Inventario_Total - promedioAñoAnt.Inventario_Total : 0;

                difencia.CostoxAnimal = promedio.CostoxAnimal != null && promedioAñoAnt.CostoxAnimal != null ? promedio.CostoxAnimal - promedioAñoAnt.CostoxAnimal : 0;
                difencia.UtilidadxAnimal = promedio.UtilidadxAnimal != null && promedioAñoAnt.UtilidadxAnimal != null ? promedio.UtilidadxAnimal - promedioAñoAnt.UtilidadxAnimal : 0;
                difencia.IngresoxAnimal = promedio.IngresoxAnimal != null && promedioAñoAnt.IngresoxAnimal != null ? promedio.IngresoxAnimal - promedioAñoAnt.IngresoxAnimal : 0;
                difencia.Porcentaje_CostoxAnimal = promedio.Porcentaje_CostoxAnimal != null && promedioAñoAnt.Porcentaje_CostoxAnimal != null ? promedio.Porcentaje_CostoxAnimal - promedioAñoAnt.Porcentaje_CostoxAnimal : 0;
                difencia.Porcentaje_UtilidadxAnimal = promedio.Porcentaje_UtilidadxAnimal != null && promedioAñoAnt.Porcentaje_UtilidadxAnimal != null ? promedio.Porcentaje_UtilidadxAnimal - promedioAñoAnt.Porcentaje_UtilidadxAnimal : 0;
                #endregion
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "DiferenciaHoja2", $"Error in {nameof(DiferenciaHoja2)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }

            return difencia;
        }

        public Hoja3 DiferenciaHoja3(Hoja3 total, Hoja3 totalAñoAnt)
        {
            Hoja3 response = new Hoja3() { Dia = "DIF #" };

            try
            {
                #region asignar datos
                response.Bp_Gem = total.Bp_Gem != null && totalAñoAnt.Bp_Gem != null ? total.Bp_Gem - totalAñoAnt.Bp_Gem : total.Bp_Gem != null ? total.Bp_Gem * (-1) : 0;
                response.Jaulas_Vivas = total.Jaulas_Vivas != null && totalAñoAnt.Jaulas_Vivas != null ? total.Jaulas_Vivas - totalAñoAnt.Jaulas_Vivas : total.Jaulas_Vivas != null ? total.Jaulas_Vivas * (-1) : 0;
                response.Jaulas_Muertas = total.Jaulas_Muertas != null && totalAñoAnt.Jaulas_Muertas != null ? total.Jaulas_Muertas - totalAñoAnt.Jaulas_Muertas : total.Jaulas_Muertas != null ? total.Jaulas_Muertas * (-1) : 0;

                response.Destete_Vivas = total.Destete_Vivas != null && totalAñoAnt.Destete_Vivas != null ? total.Destete_Vivas - totalAñoAnt.Destete_Vivas : total.Destete_Vivas != null ? total.Destete_Vivas * (-1) : 0;
                response.Destete_Muertas = total.Destete_Muertas != null && totalAñoAnt.Destete_Muertas != null ? total.Destete_Muertas - totalAñoAnt.Destete_Muertas : total.Destete_Muertas != null ? total.Destete_Muertas * (-1) : 0;

                response.Vaquillas_Muertas = total.Vaquillas_Muertas != null && totalAñoAnt.Vaquillas_Muertas != null ? total.Vaquillas_Muertas - totalAñoAnt.Vaquillas_Muertas : total.Vaquillas_Muertas != null ? total.Vaquillas_Muertas * (-1) : 0;
                response.Vaquillas_Urgente = total.Vaquillas_Urgente != null && totalAñoAnt.Vaquillas_Urgente != null ? total.Vaquillas_Urgente - totalAñoAnt.Vaquillas_Urgente : total.Vaquillas_Urgente != null ? total.Vaquillas_Urgente * (-1) : 0;
                response.Vaquillas_RF = total.Vaquillas_RF != null && totalAñoAnt.Vaquillas_RF != null ? total.Vaquillas_RF - totalAñoAnt.Vaquillas_RF : total.Vaquillas_RF != null ? total.Vaquillas_RF * (-1) : 0;
                response.Vaquillas_RR = total.Vaquillas_RR != null && totalAñoAnt.Vaquillas_RR != null ? total.Vaquillas_RR - totalAñoAnt.Vaquillas_RR : total.Vaquillas_RR != null ? total.Vaquillas_RR * (-1) : 0;
                response.Vaquillas_RG = total.Vaquillas_RG != null && totalAñoAnt.Vaquillas_RG != null ? total.Vaquillas_RG - totalAñoAnt.Vaquillas_RG : total.Vaquillas_RG != null ? total.Vaquillas_RG * (-1) : 0;
                response.Vaquillas_Otros = total.Vaquillas_Otros != null && totalAñoAnt.Vaquillas_Otros != null ? total.Vaquillas_Otros - totalAñoAnt.Vaquillas_Otros : total.Vaquillas_Otros != null ? total.Vaquillas_Otros * (-1) : 0;

                response.Vacas_Muertas = total.Vacas_Muertas != null && totalAñoAnt.Vacas_Muertas != null ? total.Vacas_Muertas - totalAñoAnt.Vacas_Muertas : total.Vacas_Muertas != null ? total.Vacas_Muertas * (-1) : 0;
                response.Vacas_Urgente = total.Vacas_Urgente != null && totalAñoAnt.Vacas_Urgente != null ? total.Vacas_Urgente - totalAñoAnt.Vacas_Urgente : total.Vacas_Urgente != null ? total.Vacas_Urgente * (-1) : 0;
                response.Vacas_RF = total.Vacas_RF != null && totalAñoAnt.Vacas_RF != null ? total.Vacas_RF - totalAñoAnt.Vacas_RF : total.Vacas_RF != null ? total.Vacas_RF * (-1) : 0;
                response.Vacas_RR = total.Vacas_RR != null && totalAñoAnt.Vacas_RR != null ? total.Vacas_RR - totalAñoAnt.Vacas_RR : total.Vacas_RR != null ? total.Vacas_RR * (-1) : 0;
                response.Vacas_RG = total.Vacas_RG != null && totalAñoAnt.Vacas_RG != null ? total.Vacas_RG - totalAñoAnt.Vacas_RG : total.Vacas_RG != null ? total.Vacas_RG * (-1) : 0;
                response.Vacas_Otros = total.Vacas_Otros != null && totalAñoAnt.Vacas_Otros != null ? total.Vacas_Otros - totalAñoAnt.Vacas_Otros : total.Vacas_Otros != null ? total.Vacas_Otros * (-1) : 0;

                response.Vaquillas_ND = total.Vaquillas_ND != null && totalAñoAnt.Vaquillas_ND != null ? total.Vaquillas_ND - totalAñoAnt.Vaquillas_ND : total.Vaquillas_ND != null ? total.Vaquillas_ND * (-1) : 0;
                response.Vaquillas_A = total.Vaquillas_A != null && totalAñoAnt.Vaquillas_A != null ? total.Vaquillas_A - totalAñoAnt.Vaquillas_A : total.Vaquillas_A != null ? total.Vaquillas_A * (-1) : 0;
                response.Vaquillas_Vivos_H = total.Vaquillas_Vivos_H != null && totalAñoAnt.Vaquillas_Vivos_H != null ? total.Vaquillas_Vivos_H - totalAñoAnt.Vaquillas_Vivos_H : total.Vaquillas_Vivos_H != null ? total.Vaquillas_Vivos_H * (-1) : 0;
                response.Vaquillas_Vivos_M = total.Vaquillas_Vivos_M != null && totalAñoAnt.Vaquillas_Vivos_M != null ? total.Vaquillas_Vivos_M - totalAñoAnt.Vaquillas_Vivos_M : total.Vaquillas_Vivos_M != null ? total.Vaquillas_Vivos_M * (-1) : 0;

                response.Vacas_ND = total.Vacas_ND != null && totalAñoAnt.Vacas_ND != null ? total.Vacas_ND - totalAñoAnt.Vacas_ND : total.Vacas_ND != null ? total.Vacas_ND * (-1) : 0;
                response.Vacas_A = total.Vacas_A != null && totalAñoAnt.Vacas_A != null ? total.Vacas_A - totalAñoAnt.Vacas_A : total.Vacas_A != null ? total.Vacas_A * (-1) : 0;
                response.Vacas_Vivos_H = total.Vacas_Vivos_H != null && totalAñoAnt.Vacas_Vivos_H != null ? total.Vacas_Vivos_H - totalAñoAnt.Vacas_Vivos_H : total.Vacas_Vivos_H != null ? total.Vacas_Vivos_H * (-1) : 0;
                response.Vacas_Vivos_M = total.Vacas_Vivos_M != null && totalAñoAnt.Vacas_Vivos_M != null ? total.Vacas_Vivos_M - totalAñoAnt.Vacas_Vivos_M : total.Vacas_Vivos_M != null ? total.Vacas_Vivos_M * (-1) : 0;

                response.SC = total.SC != null && totalAñoAnt.SC != null ? total.SC - totalAñoAnt.SC : total.SC != null ? total.SC * (-1) : 0;

                response.Partos_Vaquillas = total.Partos_Vaquillas != null && totalAñoAnt.Partos_Vaquillas != null ? total.Partos_Vaquillas - totalAñoAnt.Partos_Vaquillas : total.Partos_Vaquillas != null ? total.Partos_Vaquillas * (-1) : 0;
                response.Partos_Vacas = total.Partos_Vacas != null && totalAñoAnt.Partos_Vacas != null ? total.Partos_Vacas - totalAñoAnt.Partos_Vacas : total.Partos_Vacas != null ? total.Partos_Vacas * (-1) : 0;

                response.Muertas_Dia = total.Muertas_Dia != null && totalAñoAnt.Muertas_Dia != null ? total.Muertas_Dia - totalAñoAnt.Muertas_Dia : total.Muertas_Dia != null ? total.Muertas_Dia * (-1) : 0;
                response.Muertas_Noc = total.Muertas_Noc != null && totalAñoAnt.Muertas_Noc != null ? total.Muertas_Noc - totalAñoAnt.Muertas_Noc : total.Muertas_Noc != null ? total.Muertas_Noc * (-1) : 0;

                response.Diferencia_Calostro = total.Diferencia_Calostro != null && totalAñoAnt.Diferencia_Calostro != null ? total.Diferencia_Calostro - totalAñoAnt.Diferencia_Calostro : total.Diferencia_Calostro != null ? total.Diferencia_Calostro * (-1) : 0;
                response.Porcentaje_Calostro = total.Porcentaje_Calostro != null && totalAñoAnt.Porcentaje_Calostro != null ? total.Porcentaje_Calostro - totalAñoAnt.Porcentaje_Calostro : total.Porcentaje_Calostro != null ? total.Porcentaje_Calostro * (-1) : 0;
                response.Calidad_Calostro = total.Calidad_Calostro != null && totalAñoAnt.Calidad_Calostro != null ? total.Calidad_Calostro - totalAñoAnt.Calidad_Calostro : total.Calidad_Calostro != null ? total.Calidad_Calostro * (-1) : 0;
                #endregion
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "DiferenciaHoja3", $"Error in {nameof(DiferenciaHoja3)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }

            return response;
        }

        public Hoja4 DiferenciaHoja4(Hoja4 total, Hoja4 totalAñoAnt)
        {
            Hoja4 response = new Hoja4();
            response.Dia = "DIF #";

            try
            {
                response.Ubre_MA = total != null && totalAñoAnt != null ? total.Ubre_MA - totalAñoAnt.Ubre_MA : 0;
                response.Ubre_SL = total != null && totalAñoAnt != null ? total.Ubre_SL - totalAñoAnt.Ubre_SL : 0;
                response.Metabolicos_FL = total != null && totalAñoAnt != null ? total.Metabolicos_FL - totalAñoAnt.Metabolicos_FL : 0;
                response.Metabolicos_CET = total != null && totalAñoAnt != null ? total.Metabolicos_CET - totalAñoAnt.Metabolicos_CET : 0;
                response.Locomotores_BE = total != null && totalAñoAnt != null ? total.Locomotores_BE - totalAñoAnt.Locomotores_BE : 0;
                response.Locomotores_TRA = total != null && totalAñoAnt != null ? total.Locomotores_TRA - totalAñoAnt.Locomotores_TRA : 0;
                response.Locomotores_GA = total != null && totalAñoAnt != null ? total.Locomotores_GA - totalAñoAnt.Locomotores_GA : 0;
                response.Digestivos_AC = total != null && totalAñoAnt != null ? total.Digestivos_AC - totalAñoAnt.Digestivos_AC : 0;
                response.Digestivos_ES = total != null && totalAñoAnt != null ? total.Digestivos_ES - totalAñoAnt.Digestivos_ES : 0;
                response.Digestivos_DI = total != null && totalAñoAnt != null ? total.Digestivos_DI - totalAñoAnt.Digestivos_DI : 0;
                response.Digestivos_TI = total != null && totalAñoAnt != null ? total.Digestivos_TI - totalAñoAnt.Digestivos_TI : 0;
                response.Reproductivos_RE = total != null && totalAñoAnt != null ? total.Reproductivos_RE - totalAñoAnt.Reproductivos_RE : 0;
                response.Reproductivos_ME = total != null && totalAñoAnt != null ? total.Reproductivos_ME - totalAñoAnt.Reproductivos_ME : 0;
                response.Reproductivos_PIO = total != null && totalAñoAnt != null ? total.Reproductivos_PIO - totalAñoAnt.Reproductivos_PIO : 0;
                response.Reproductivos_QUI = total != null && totalAñoAnt != null ? total.Reproductivos_QUI - totalAñoAnt.Reproductivos_QUI : 0;
                response.Reproductivos_CS = total != null && totalAñoAnt != null ? total.Reproductivos_CS - totalAñoAnt.Reproductivos_CS : 0;
                response.Respiratorios_Neu = total != null && totalAñoAnt != null ? total.Respiratorios_Neu - totalAñoAnt.Respiratorios_Neu : 0;
                response.Becerras_Neu = total != null && totalAñoAnt != null ? total.Becerras_Neu - totalAñoAnt.Becerras_Neu : 0;
                response.Becerras_Fie = total != null && totalAñoAnt != null ? total.Becerras_Fie - totalAñoAnt.Becerras_Fie : 0;
                response.Becerras_Di = total != null && totalAñoAnt != null ? total.Becerras_Di - totalAñoAnt.Becerras_Di : 0;
                response.Becerras_Conj = total != null && totalAñoAnt != null ? total.Becerras_Conj - totalAñoAnt.Becerras_Conj : 0;
                response.Vacas_Diag = total != null && totalAñoAnt != null ? total.Vacas_Diag - totalAñoAnt.Vacas_Diag : 0;
                response.Vacas_Pren = total != null && totalAñoAnt != null ? total.Vacas_Pren - totalAñoAnt.Vacas_Pren : 0;
                response.Vacas_Porcentaje_Pren = total != null && totalAñoAnt != null ? total.Vacas_Porcentaje_Pren - totalAñoAnt.Vacas_Porcentaje_Pren : 0;
                response.Vacas_Vacias = total != null && totalAñoAnt != null ? total.Vacas_Vacias - totalAñoAnt.Vacas_Vacias : 0;
                response.Vacas_Porcentaje_Vacias = total != null && totalAñoAnt != null ? total.Vacas_Porcentaje_Vacias - totalAñoAnt.Vacas_Porcentaje_Vacias : 0;
                response.Vaquillas_Diag = total != null && totalAñoAnt != null ? total.Vaquillas_Diag - totalAñoAnt.Vaquillas_Diag : 0;
                response.Vaquillas_Pren = total != null && totalAñoAnt != null ? total.Vaquillas_Pren - totalAñoAnt.Vaquillas_Pren : 0;
                response.Vaquillas_Porcentaje_Pren = total != null && totalAñoAnt != null ? total.Vaquillas_Porcentaje_Pren - totalAñoAnt.Vaquillas_Porcentaje_Pren : 0;
                response.Vaquillas_Vacias = total != null && totalAñoAnt != null ? total.Vaquillas_Vacias - totalAñoAnt.Vaquillas_Vacias : 0;
                response.Vaquillas_Porcentaje_Vacias = total != null && totalAñoAnt != null ? total.Vaquillas_Porcentaje_Vacias - totalAñoAnt.Vaquillas_Porcentaje_Vacias : 0;
                response.Abortos_Vaquillas = total != null && totalAñoAnt != null ? total.Abortos_Vaquillas - totalAñoAnt.Abortos_Vaquillas : 0;
                response.Abortos_Vacas = total != null && totalAñoAnt != null ? total.Abortos_Vacas - totalAñoAnt.Abortos_Vacas : 0;
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "DiferenciaHoja4", $"Error in {nameof(DiferenciaHoja4)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }

            return response;
        }
        #endregion

        #region porcentaje diferencias
        public Hoja1 PorcentajeDiferenciaHoja1(Hoja1 diferencia, Hoja1 añoAnterior)
        {
            Hoja1 porcentaje = new Hoja1() { Dia = "DIF %" };

            try
            {
                porcentaje.Ordeño = diferencia.Ordeño != null && añoAnterior.Ordeño != null && añoAnterior.Ordeño != 0 ? diferencia.Ordeño / añoAnterior.Ordeño : 0;
                porcentaje.Secas = diferencia.Secas != null && añoAnterior.Secas != null && añoAnterior.Secas != 0 ? diferencia.Secas / añoAnterior.Secas : 0;
                porcentaje.Hato = diferencia.Hato != null && añoAnterior.Hato != null && añoAnterior.Hato != 0 ? diferencia.Hato / añoAnterior.Hato : 0;
                porcentaje.Porcentaje_Lact = diferencia.Porcentaje_Lact != null && añoAnterior.Porcentaje_Lact != null && añoAnterior.Porcentaje_Lact != 0 ? diferencia.Porcentaje_Lact / añoAnterior.Porcentaje_Lact : 0;
                porcentaje.Porcentaje_Prot = diferencia.Porcentaje_Prot != null && añoAnterior.Porcentaje_Prot != null && añoAnterior.Porcentaje_Prot != 0 ? diferencia.Porcentaje_Prot / añoAnterior.Porcentaje_Prot : 0;
                porcentaje.Urea = diferencia.Urea != null && añoAnterior.Urea != null && añoAnterior.Urea != 0 ? diferencia.Urea / añoAnterior.Urea : 0;
                porcentaje.Porcentaje_Grasa = diferencia.Porcentaje_Grasa != null && añoAnterior.Porcentaje_Grasa != null && añoAnterior.Porcentaje_Grasa != 0 ? diferencia.Porcentaje_Grasa / añoAnterior.Porcentaje_Grasa : 0;
                porcentaje.CCS = diferencia.CCS != null && añoAnterior.CCS != null && añoAnterior.CCS != 0 ? diferencia.CCS / añoAnterior.CCS : 0;
                porcentaje.CTD = diferencia.CTD != null && añoAnterior.CTD != null && añoAnterior.CTD != 0 ? diferencia.CTD / añoAnterior.CTD : 0;
                porcentaje.DEC = null;

                porcentaje.Leche = diferencia.Leche != null && añoAnterior.Leche != null && añoAnterior.Leche != 0 ? diferencia.Leche / añoAnterior.Leche : 0;
                porcentaje.Antib = diferencia.Antib != null && añoAnterior.Antib != null && añoAnterior.Antib != 0 ? diferencia.Antib / añoAnterior.Antib : 0;
                porcentaje.Media = diferencia.Media != null && añoAnterior.Media != null && añoAnterior.Media != 0 ? diferencia.Media / añoAnterior.Media : 0;
                porcentaje.Total = diferencia.Total != null && añoAnterior.Total != null && añoAnterior.Total != 0 ? diferencia.Total / añoAnterior.Total : 0;
                porcentaje.DEL = diferencia.DEL != null && añoAnterior.DEL != null && añoAnterior.DEL != 0 ? diferencia.DEL / añoAnterior.DEL : 0;
                porcentaje.ANT = diferencia.ANT != null && añoAnterior.ANT != null && añoAnterior.ANT != 0 ? diferencia.ANT / añoAnterior.ANT : 0;

                porcentaje.EA = diferencia.EA != null && añoAnterior.EA != null && añoAnterior.EA != 0 ? diferencia.EA / añoAnterior.EA : 0;
                porcentaje.ILCA = diferencia.ILCA != null && añoAnterior.ILCA != null && añoAnterior.ILCA != 0 ? diferencia.ILCA / añoAnterior.ILCA : 0;
                porcentaje.IC = diferencia.IC != null && añoAnterior.IC != null && añoAnterior.IC != 0 ? diferencia.IC / añoAnterior.IC : 0;
                porcentaje.Costo_Litro = diferencia.Costo_Litro != null && añoAnterior.Costo_Litro != null && añoAnterior.Costo_Litro != 0 ? diferencia.Costo_Litro / añoAnterior.Costo_Litro : 0;
                porcentaje.MH = diferencia.MH != null && añoAnterior.MH != null && añoAnterior.MH != 0 ? diferencia.MH / añoAnterior.MH : 0;
                porcentaje.Porcentaje_MS = diferencia.Porcentaje_MS != null && añoAnterior.Porcentaje_MS != null && añoAnterior.Porcentaje_MS != 0 ? diferencia.Porcentaje_MS / añoAnterior.Porcentaje_MS : 0;
                porcentaje.MS = diferencia.MS != null && añoAnterior.MS != null && añoAnterior.MS != 0 ? diferencia.MS / añoAnterior.MS : 0;
                porcentaje.SA = diferencia.SA != null && añoAnterior.SA != null && añoAnterior.SA != 0 ? diferencia.SA / añoAnterior.SA : 0;
                porcentaje.MSS = diferencia.MSS != null && añoAnterior.MSS != null && añoAnterior.MSS != 0 ? diferencia.MSS / añoAnterior.MSS : 0;
                porcentaje.EAS = diferencia.EAS != null && añoAnterior.EAS != null && añoAnterior.EAS != 0 ? diferencia.EAS / añoAnterior.EAS : 0;
                porcentaje.Porcentaje_Sob = diferencia.Porcentaje_Sob != null && añoAnterior.Porcentaje_Sob != null && añoAnterior.Porcentaje_Sob != 0 ? diferencia.Porcentaje_Sob / añoAnterior.Porcentaje_Sob : 0;
                porcentaje.Costo_Prod = diferencia.Costo_Prod != null && añoAnterior.Costo_Prod != null && añoAnterior.Costo_Prod != 0 ? diferencia.Costo_Prod / añoAnterior.Costo_Prod : 0;
                porcentaje.Costo_MS = diferencia.Costo_MS != null && añoAnterior.Costo_MS != null && añoAnterior.Costo_MS != 0 ? diferencia.Costo_MS / añoAnterior.Costo_MS : 0;

                porcentaje.Cribas_N1 = diferencia.Cribas_N1 != null && añoAnterior.Cribas_N1 != null && añoAnterior.Cribas_N1 != 0 ? diferencia.Cribas_N1 / añoAnterior.Cribas_N1 : 0;
                porcentaje.Cribas_N2 = diferencia.Cribas_N2 != null && añoAnterior.Cribas_N2 != null && añoAnterior.Cribas_N2 != 0 ? diferencia.Cribas_N2 / añoAnterior.Cribas_N2 : 0;
                porcentaje.Cribas_N3 = diferencia.Cribas_N3 != null && añoAnterior.Cribas_N3 != null && añoAnterior.Cribas_N3 != 0 ? diferencia.Cribas_N3 / añoAnterior.Cribas_N3 : 0;
                porcentaje.Cribas_N4 = diferencia.Cribas_N4 != null && añoAnterior.Cribas_N4 != null && añoAnterior.Cribas_N4 != 0 ? diferencia.Cribas_N4 / añoAnterior.Cribas_N4 : 0;

                porcentaje.NoID_Ses1 = diferencia.NoID_Ses1 != null && añoAnterior.NoID_Ses1 != null && añoAnterior.NoID_Ses1 != 0 ? diferencia.NoID_Ses1 / añoAnterior.NoID_Ses1 : 0;
                porcentaje.NoID_Ses2 = diferencia.NoID_Ses2 != null && añoAnterior.NoID_Ses2 != null && añoAnterior.NoID_Ses2 != 0 ? diferencia.NoID_Ses2 / añoAnterior.NoID_Ses2 : 0;
                porcentaje.NoID_Ses3 = diferencia.NoID_Ses3 != null && añoAnterior.NoID_Ses3 != null && añoAnterior.NoID_Ses3 != 0 ? diferencia.NoID_Ses3 / añoAnterior.NoID_Ses3 : 0;

                porcentaje.SES1 = diferencia.SES1 != null && añoAnterior.SES1 != null && añoAnterior.SES1 != 0 ? diferencia.SES1 / añoAnterior.SES1 : 0;
                porcentaje.SES2 = diferencia.SES2 != null && añoAnterior.SES2 != null && añoAnterior.SES2 != 0 ? diferencia.SES2 / añoAnterior.SES2 : 0;
                porcentaje.SES3 = diferencia.SES3 != null && añoAnterior.SES3 != null && añoAnterior.SES3 != 0 ? diferencia.SES3 / añoAnterior.SES3 : 0;

                diferencia.Porcentaje_Revueltas = diferencia.Porcentaje_Revueltas != null && añoAnterior.Porcentaje_Revueltas != null && añoAnterior.Porcentaje_Revueltas != 0 ? diferencia.Porcentaje_Revueltas / añoAnterior.Porcentaje_Revueltas : 0;
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "PorcentajeDiferenciaHoja1", $"Error in {nameof(PorcentajeDiferenciaHoja1)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }

            return porcentaje;
        }

        public Hoja2 PorcentajeDiferenciaHoja2(Hoja2 diferencia, Hoja2 promedioAñoAnt)
        {
            Hoja2 porcentaje = new Hoja2() { Dia = "DIF %" };

            try
            {
                #region Crianza 
                porcentaje.Crianza_Inventario = diferencia.Crianza_Inventario != null && promedioAñoAnt.Crianza_Inventario != null && promedioAñoAnt.Crianza_Inventario != 0 ? diferencia.Crianza_Inventario / promedioAñoAnt.Crianza_Inventario : 0;
                #endregion

                #region Jaulas 
                porcentaje.Jaulas_Inventario = diferencia.Jaulas_Inventario != null && promedioAñoAnt.Jaulas_Inventario != null && promedioAñoAnt.Jaulas_Inventario != 0 ? diferencia.Jaulas_Inventario / promedioAñoAnt.Jaulas_Inventario : 0;
                porcentaje.Jaulas_Costo = diferencia.Jaulas_Inventario != null && promedioAñoAnt.Jaulas_Inventario != null && promedioAñoAnt.Jaulas_Inventario != 0 ? diferencia.Jaulas_Inventario / promedioAñoAnt.Jaulas_Inventario : 0;
                #endregion

                #region Crecimiento 
                porcentaje.Crecimiento_Inventario = diferencia.Crecimiento_Inventario != null && promedioAñoAnt.Crecimiento_Inventario != null && promedioAñoAnt.Crecimiento_Inventario != 0 ? diferencia.Crecimiento_Inventario / promedioAñoAnt.Crecimiento_Inventario : 0;
                porcentaje.Crecimiento_MH = diferencia.Crecimiento_MH != null && promedioAñoAnt.Crecimiento_MH != null && promedioAñoAnt.Crecimiento_MH != 0 ? diferencia.Crecimiento_MH / promedioAñoAnt.Crecimiento_MH : 0;
                porcentaje.Crecimiento_Costo = diferencia.Crecimiento_Costo != null && promedioAñoAnt.Crecimiento_Costo != null && promedioAñoAnt.Crecimiento_Costo != 0 ? diferencia.Crecimiento_Costo / promedioAñoAnt.Jaulas_Inventario : 0;
                porcentaje.Crecimiento_PorcentajeMS = diferencia.Crecimiento_PorcentajeMS != null && promedioAñoAnt.Crecimiento_PorcentajeMS != null && promedioAñoAnt.Crecimiento_PorcentajeMS != 0 ? diferencia.Crecimiento_PorcentajeMS / promedioAñoAnt.Crecimiento_PorcentajeMS : 0;
                porcentaje.Crecimiento_MS = diferencia.Crecimiento_MS != null && promedioAñoAnt.Crecimiento_MS != null && promedioAñoAnt.Crecimiento_MS != 0 ? diferencia.Crecimiento_MS / promedioAñoAnt.Crecimiento_MS : 0;
                porcentaje.Crecimiento_CostoMS = diferencia.Crecimiento_CostoMS != null && promedioAñoAnt.Crecimiento_CostoMS != null && promedioAñoAnt.Crecimiento_CostoMS != 0 ? diferencia.Crecimiento_CostoMS / promedioAñoAnt.Crecimiento_CostoMS : 0;
                #endregion

                #region Desarrollo 
                porcentaje.Desarrollo_Inventario = diferencia.Desarrollo_Inventario != null && promedioAñoAnt.Desarrollo_Inventario != null && promedioAñoAnt.Desarrollo_Inventario != 0 ? diferencia.Desarrollo_Inventario / promedioAñoAnt.Desarrollo_Inventario : 0;
                porcentaje.Desarrollo_MH = diferencia.Desarrollo_MH != null && promedioAñoAnt.Desarrollo_MH != null && promedioAñoAnt.Desarrollo_MH != 0 ? diferencia.Desarrollo_MH / promedioAñoAnt.Desarrollo_MH : 0;
                porcentaje.Desarrollo_Costo = diferencia.Desarrollo_Costo != null && promedioAñoAnt.Desarrollo_Costo != null && promedioAñoAnt.Desarrollo_Costo != 0 ? diferencia.Desarrollo_Costo / promedioAñoAnt.Desarrollo_Costo : 0;
                porcentaje.Desarrollo_PorcentajeMS = diferencia.Desarrollo_PorcentajeMS != null && promedioAñoAnt.Desarrollo_PorcentajeMS != null && promedioAñoAnt.Desarrollo_PorcentajeMS != 0 ? diferencia.Desarrollo_PorcentajeMS / promedioAñoAnt.Desarrollo_PorcentajeMS : 0;
                porcentaje.Desarrollo_MS = diferencia.Desarrollo_MS != null && promedioAñoAnt.Desarrollo_MS != null && promedioAñoAnt.Desarrollo_MS != 0 ? diferencia.Desarrollo_MS / promedioAñoAnt.Desarrollo_MS : 0;
                porcentaje.Desarrollo_CostoMS = diferencia.Desarrollo_CostoMS != null && promedioAñoAnt.Desarrollo_CostoMS != null && promedioAñoAnt.Desarrollo_CostoMS != 0 ? diferencia.Desarrollo_CostoMS / promedioAñoAnt.Desarrollo_CostoMS : 0;
                #endregion

                #region Vaquillas 
                porcentaje.Vaquillas_Inventario = diferencia.Vaquillas_Inventario != null && promedioAñoAnt.Vaquillas_Inventario != null && promedioAñoAnt.Vaquillas_Inventario != 0 ? diferencia.Vaquillas_Inventario / promedioAñoAnt.Vaquillas_Inventario : 0;
                porcentaje.Vaquillas_MH = diferencia.Vaquillas_MH != null && promedioAñoAnt.Vaquillas_MH != null && promedioAñoAnt.Vaquillas_MH != 0 ? diferencia.Vaquillas_MH / promedioAñoAnt.Vaquillas_MH : 0;
                porcentaje.Vaquillas_Costo = diferencia.Vaquillas_Costo != null && promedioAñoAnt.Vaquillas_Costo != null && promedioAñoAnt.Vaquillas_Costo != 0 ? diferencia.Vaquillas_Costo / promedioAñoAnt.Vaquillas_Costo : 0;
                porcentaje.Vaquillas_PorcentajeMS = diferencia.Vaquillas_PorcentajeMS != null && promedioAñoAnt.Vaquillas_PorcentajeMS != null && promedioAñoAnt.Vaquillas_PorcentajeMS != 0 ? diferencia.Vaquillas_PorcentajeMS / promedioAñoAnt.Vaquillas_PorcentajeMS : 0;
                porcentaje.Vaquillas_MS = diferencia.Vaquillas_MS != null && promedioAñoAnt.Vaquillas_MS != null && promedioAñoAnt.Vaquillas_MS != 0 ? diferencia.Vaquillas_MS / promedioAñoAnt.Vaquillas_MS : 0;
                porcentaje.Vaquillas_CostoMS = diferencia.Vaquillas_CostoMS != null && promedioAñoAnt.Vaquillas_CostoMS != null && promedioAñoAnt.Vaquillas_CostoMS != 0 ? diferencia.Vaquillas_CostoMS / promedioAñoAnt.Vaquillas_CostoMS : 0;
                #endregion

                #region Secas 
                porcentaje.Secas_Inventario = diferencia.Secas_Inventario != null && promedioAñoAnt.Secas_Inventario != null && promedioAñoAnt.Secas_Inventario != 0 ? diferencia.Secas_Inventario / promedioAñoAnt.Secas_Inventario : 0;
                porcentaje.Secas_MH = diferencia.Secas_MH != null && promedioAñoAnt.Secas_MH != null && promedioAñoAnt.Secas_MH != 0 ? diferencia.Secas_MH / promedioAñoAnt.Secas_MH : 0;
                porcentaje.Secas_PorcentajeMS = diferencia.Secas_PorcentajeMS != null && promedioAñoAnt.Secas_PorcentajeMS != null && promedioAñoAnt.Secas_PorcentajeMS != 0 ? diferencia.Secas_PorcentajeMS / promedioAñoAnt.Secas_PorcentajeMS : 0;
                porcentaje.Secas_MS = diferencia.Secas_MS != null && promedioAñoAnt.Secas_MS != null && promedioAñoAnt.Secas_MS != 0 ? diferencia.Secas_MS / promedioAñoAnt.Secas_MS : 0;
                porcentaje.Secas_SA = diferencia.Secas_SA != null && promedioAñoAnt.Secas_SA != null && promedioAñoAnt.Secas_SA != 0 ? diferencia.Secas_SA / promedioAñoAnt.Secas_SA : 0;
                porcentaje.Secas_Mss = diferencia.Secas_Mss != null && promedioAñoAnt.Secas_Mss != null && promedioAñoAnt.Secas_Mss != 0 ? diferencia.Secas_Mss / promedioAñoAnt.Secas_Mss : 0;
                porcentaje.Secas_PorcentajeSob = diferencia.Secas_PorcentajeSob != null && promedioAñoAnt.Secas_PorcentajeSob != null && promedioAñoAnt.Secas_PorcentajeSob != 0 ? diferencia.Secas_PorcentajeSob / promedioAñoAnt.Secas_PorcentajeSob : 0;
                porcentaje.Secas_Costo = diferencia.Secas_Costo != null && promedioAñoAnt.Secas_Costo != null && promedioAñoAnt.Secas_Costo != 0 ? diferencia.Secas_Costo / promedioAñoAnt.Secas_Costo : 0;
                porcentaje.Secas_CostoMS = diferencia.Secas_CostoMS != null && promedioAñoAnt.Secas_CostoMS != null && promedioAñoAnt.Secas_CostoMS != 0 ? diferencia.Secas_CostoMS / promedioAñoAnt.Secas_CostoMS : 0;
                #endregion

                #region Reto 
                porcentaje.Reto_Inventario = diferencia.Reto_Inventario != null && promedioAñoAnt.Reto_Inventario != null && promedioAñoAnt.Reto_Inventario != 0 ? diferencia.Reto_Inventario / promedioAñoAnt.Reto_Inventario : 0;
                porcentaje.Reto_MH = diferencia.Reto_MH != null && promedioAñoAnt.Reto_MH != null && promedioAñoAnt.Reto_MH != 0 ? diferencia.Reto_MH / promedioAñoAnt.Reto_MH : 0;
                porcentaje.Reto_PorcentajeMS = diferencia.Reto_PorcentajeMS != null && promedioAñoAnt.Reto_PorcentajeMS != null && promedioAñoAnt.Reto_PorcentajeMS != 0 ? diferencia.Reto_PorcentajeMS / promedioAñoAnt.Reto_PorcentajeMS : 0;
                porcentaje.Reto_MS = diferencia.Reto_MS != null && promedioAñoAnt.Reto_MS != null && promedioAñoAnt.Reto_MS != 0 ? diferencia.Reto_MS / promedioAñoAnt.Reto_MS : 0;
                porcentaje.Reto_SA = diferencia.Reto_SA != null && promedioAñoAnt.Reto_SA != null && promedioAñoAnt.Reto_SA != 0 ? diferencia.Reto_SA / promedioAñoAnt.Reto_SA : 0;
                porcentaje.Reto_Mss = diferencia.Reto_Mss != null && promedioAñoAnt.Reto_Mss != null && promedioAñoAnt.Reto_Mss != 0 ? diferencia.Reto_Mss / promedioAñoAnt.Reto_Mss : 0;
                porcentaje.Reto_PorcentajeSob = diferencia.Reto_PorcentajeSob != null && promedioAñoAnt.Reto_PorcentajeSob != null && promedioAñoAnt.Reto_PorcentajeSob != 0 ? diferencia.Reto_PorcentajeSob / promedioAñoAnt.Reto_PorcentajeSob : 0;
                porcentaje.Reto_Costo = diferencia.Reto_Costo != null && promedioAñoAnt.Reto_Costo != null && promedioAñoAnt.Reto_Costo != 0 ? diferencia.Reto_Costo / promedioAñoAnt.Reto_Costo : 0;
                porcentaje.Reto_CostoMS = diferencia.Reto_CostoMS != null && promedioAñoAnt.Reto_CostoMS != null && promedioAñoAnt.Reto_CostoMS != 0 ? diferencia.Reto_CostoMS / promedioAñoAnt.Reto_CostoMS : 0;
                #endregion

                #region Utilidad 
                porcentaje.Inventario_Total = diferencia.Inventario_Total != null && promedioAñoAnt.Inventario_Total != null && promedioAñoAnt.Inventario_Total != 0 ? diferencia.Inventario_Total / promedioAñoAnt.Inventario_Total : 0;

                porcentaje.CostoxAnimal = diferencia.CostoxAnimal != null && promedioAñoAnt.CostoxAnimal != null && promedioAñoAnt.CostoxAnimal != 0 ? diferencia.CostoxAnimal / promedioAñoAnt.CostoxAnimal : 0;

                porcentaje.UtilidadxAnimal = diferencia.UtilidadxAnimal != null && promedioAñoAnt.UtilidadxAnimal != null && promedioAñoAnt.UtilidadxAnimal != 0 ? diferencia.UtilidadxAnimal / promedioAñoAnt.UtilidadxAnimal : 0;

                porcentaje.IngresoxAnimal = diferencia.IngresoxAnimal != null && promedioAñoAnt.IngresoxAnimal != null && promedioAñoAnt.IngresoxAnimal != 0 ? diferencia.IngresoxAnimal / promedioAñoAnt.IngresoxAnimal : 0;

                porcentaje.Porcentaje_CostoxAnimal = diferencia.Porcentaje_CostoxAnimal != null && promedioAñoAnt.Porcentaje_CostoxAnimal != null && promedioAñoAnt.Porcentaje_CostoxAnimal != 0 ? diferencia.Porcentaje_CostoxAnimal / promedioAñoAnt.Porcentaje_CostoxAnimal : 0;

                porcentaje.Porcentaje_UtilidadxAnimal = diferencia.Porcentaje_UtilidadxAnimal != null && promedioAñoAnt.Porcentaje_UtilidadxAnimal != null && promedioAñoAnt.Porcentaje_UtilidadxAnimal != 0 ? diferencia.Porcentaje_UtilidadxAnimal / promedioAñoAnt.Porcentaje_UtilidadxAnimal : 0;
                #endregion
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "PorcentajeDiferenciaHoja2", $"Error in {nameof(PorcentajeDiferenciaHoja2)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }

            return porcentaje;

        }

        public Hoja3 PorcentajeDiferenciaReporteHoja3(Hoja3 diferencia, Hoja3 totalAñoAnt)
        {
            Hoja3 response = new Hoja3() { Dia = "DIF %" };

            try
            {
                #region asignar datos
                response.Bp_Gem = diferencia.Bp_Gem != null && totalAñoAnt.Bp_Gem != null && totalAñoAnt.Bp_Gem != 0 ? diferencia.Bp_Gem / totalAñoAnt.Bp_Gem : 0;
                response.Jaulas_Vivas = diferencia.Jaulas_Vivas != null && totalAñoAnt.Jaulas_Vivas != null && totalAñoAnt.Jaulas_Vivas != 0 ? diferencia.Jaulas_Vivas / totalAñoAnt.Jaulas_Vivas : 0;
                response.Jaulas_Muertas = diferencia.Jaulas_Muertas != null && totalAñoAnt.Jaulas_Muertas != null && totalAñoAnt.Jaulas_Muertas != 0 ? diferencia.Jaulas_Muertas / totalAñoAnt.Jaulas_Muertas : 0;

                response.Destete_Vivas = diferencia.Destete_Vivas != null && totalAñoAnt.Destete_Vivas != null && totalAñoAnt.Destete_Vivas != 0 ? diferencia.Destete_Vivas / totalAñoAnt.Destete_Vivas : 0;
                response.Destete_Muertas = diferencia.Destete_Muertas != null && totalAñoAnt.Destete_Muertas != null && totalAñoAnt.Destete_Muertas != 0 ? diferencia.Destete_Muertas / totalAñoAnt.Destete_Muertas : 0;

                response.Vaquillas_Muertas = diferencia.Vaquillas_Muertas != null && totalAñoAnt.Vaquillas_Muertas != null && totalAñoAnt.Vaquillas_Muertas != 0 ? diferencia.Vaquillas_Muertas / totalAñoAnt.Vaquillas_Muertas : 0;
                response.Vaquillas_Urgente = diferencia.Vaquillas_Urgente != null && totalAñoAnt.Vaquillas_Urgente != null && totalAñoAnt.Vaquillas_Urgente != 0 ? diferencia.Vaquillas_Urgente / totalAñoAnt.Vaquillas_Urgente : 0;
                response.Vaquillas_RF = diferencia.Vaquillas_RF != null && totalAñoAnt.Vaquillas_RF != null && totalAñoAnt.Vaquillas_RF != 0 ? diferencia.Vaquillas_RF / totalAñoAnt.Vaquillas_RF : 0;
                response.Vaquillas_RR = diferencia.Vaquillas_RR != null && totalAñoAnt.Vaquillas_RR != null && totalAñoAnt.Vaquillas_RR != 0 ? diferencia.Vaquillas_RR / totalAñoAnt.Vaquillas_RR : 0;
                response.Vaquillas_RG = diferencia.Vaquillas_RG != null && totalAñoAnt.Vaquillas_RG != null && totalAñoAnt.Vaquillas_RG != 0 ? diferencia.Vaquillas_RG / totalAñoAnt.Vaquillas_RG : 0;
                response.Vaquillas_Otros = diferencia.Vaquillas_Otros != null && totalAñoAnt.Vaquillas_Otros != null && totalAñoAnt.Vaquillas_Otros != 0 ? diferencia.Vaquillas_Otros / totalAñoAnt.Vaquillas_Otros : 0;

                response.Vacas_Muertas = diferencia.Vacas_Muertas != null && totalAñoAnt.Vacas_Muertas != null && totalAñoAnt.Vacas_Muertas != 0 ? diferencia.Vacas_Muertas / totalAñoAnt.Vacas_Muertas : 0;
                response.Vacas_Urgente = diferencia.Vacas_Urgente != null && totalAñoAnt.Vacas_Urgente != null && totalAñoAnt.Vacas_Urgente != 0 ? diferencia.Vacas_Urgente / totalAñoAnt.Vacas_Urgente : 0;
                response.Vacas_RF = diferencia.Vacas_RF != null && totalAñoAnt.Vacas_RF != null && totalAñoAnt.Vacas_RF != 0 ? diferencia.Vacas_RF / totalAñoAnt.Vacas_RF : 0;
                response.Vacas_RR = diferencia.Vacas_RR != null && totalAñoAnt.Vacas_RR != null && totalAñoAnt.Vacas_RR != 0 ? diferencia.Vacas_RR / totalAñoAnt.Vacas_RR : 0;
                response.Vacas_RG = diferencia.Vacas_RG != null && totalAñoAnt.Vacas_RG != null && totalAñoAnt.Vacas_RG != 0 ? diferencia.Vacas_RG / totalAñoAnt.Vacas_RG : 0;
                response.Vacas_Otros = diferencia.Vacas_Otros != null && totalAñoAnt.Vacas_Otros != null && totalAñoAnt.Vacas_Otros != 0 ? diferencia.Vacas_Otros / totalAñoAnt.Vacas_Otros : 0;

                response.Vaquillas_ND = diferencia.Vaquillas_ND != null && totalAñoAnt.Vaquillas_ND != null && totalAñoAnt.Vaquillas_ND != 0 ? diferencia.Vaquillas_ND / totalAñoAnt.Vaquillas_ND : 0;
                response.Vaquillas_A = diferencia.Vaquillas_A != null && totalAñoAnt.Vaquillas_A != null && totalAñoAnt.Vaquillas_A != 0 ? diferencia.Vaquillas_A / totalAñoAnt.Vaquillas_A : 0;
                response.Vaquillas_Vivos_H = diferencia.Vaquillas_Vivos_H != null && totalAñoAnt.Vaquillas_Vivos_H != null && totalAñoAnt.Vaquillas_Vivos_H != 0 ? diferencia.Vaquillas_Vivos_H / totalAñoAnt.Vaquillas_Vivos_H : 0;
                response.Vaquillas_Vivos_M = diferencia.Vaquillas_Vivos_M != null && totalAñoAnt.Vaquillas_Vivos_M != null && totalAñoAnt.Vaquillas_Vivos_M != 0 ? diferencia.Vaquillas_Vivos_M / totalAñoAnt.Vaquillas_Vivos_M : 0;

                response.Vacas_ND = diferencia.Vacas_ND != null && totalAñoAnt.Vacas_ND != null && totalAñoAnt.Vacas_ND != 0 ? diferencia.Vacas_ND / totalAñoAnt.Vacas_ND : 0;
                response.Vacas_A = diferencia.Vacas_A != null && totalAñoAnt.Vacas_A != null && totalAñoAnt.Vacas_A != 0 ? diferencia.Vacas_A / totalAñoAnt.Vacas_A : 0;
                response.Vacas_Vivos_H = diferencia.Vacas_Vivos_H != null && totalAñoAnt.Vacas_Vivos_H != null && totalAñoAnt.Vacas_Vivos_H != 0 ? diferencia.Vacas_Vivos_H / totalAñoAnt.Vacas_Vivos_H : 0;
                response.Vacas_Vivos_M = diferencia.Vacas_Vivos_M != null && totalAñoAnt.Vacas_Vivos_M != null && totalAñoAnt.Vacas_Vivos_M != 0 ? diferencia.Vacas_Vivos_M / totalAñoAnt.Vacas_Vivos_M : 0;

                response.SC = diferencia.SC != null && totalAñoAnt.SC != null && totalAñoAnt.SC != 0 ? diferencia.SC / totalAñoAnt.SC : 0;

                response.Partos_Vaquillas = diferencia.Partos_Vaquillas != null && totalAñoAnt.Partos_Vaquillas != null && totalAñoAnt.Partos_Vaquillas != 0 ? diferencia.Partos_Vaquillas / totalAñoAnt.Partos_Vaquillas : 0;
                response.Partos_Vacas = diferencia.Partos_Vacas != null && totalAñoAnt.Partos_Vacas != null && totalAñoAnt.Partos_Vacas != 0 ? diferencia.Partos_Vacas / totalAñoAnt.Partos_Vacas : 0;

                response.Muertas_Dia = diferencia.Muertas_Dia != null && totalAñoAnt.Muertas_Dia != null && totalAñoAnt.Muertas_Dia != 0 ? diferencia.Muertas_Dia / totalAñoAnt.Muertas_Dia : 0;
                response.Muertas_Noc = diferencia.Muertas_Noc != null && totalAñoAnt.Muertas_Noc != null && totalAñoAnt.Muertas_Noc != 0 ? diferencia.Muertas_Noc / totalAñoAnt.Muertas_Noc : 0;

                response.Diferencia_Calostro = diferencia.Diferencia_Calostro != null && totalAñoAnt.Diferencia_Calostro != null && totalAñoAnt.Diferencia_Calostro != 0 ? diferencia.Diferencia_Calostro / totalAñoAnt.Diferencia_Calostro : 0;
                response.Porcentaje_Calostro = diferencia.Porcentaje_Calostro != null && totalAñoAnt.Porcentaje_Calostro != null && totalAñoAnt.Porcentaje_Calostro != 0 ? diferencia.Porcentaje_Calostro / totalAñoAnt.Porcentaje_Calostro : 0;
                response.Calidad_Calostro = diferencia.Calidad_Calostro != null && totalAñoAnt.Calidad_Calostro != null && totalAñoAnt.Calidad_Calostro != 0 ? diferencia.Calidad_Calostro / totalAñoAnt.Calidad_Calostro : 0;
                #endregion
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "PorcentajeDiferenciaReporteHoja3", $"Error in {nameof(PorcentajeDiferenciaReporteHoja3)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }

            return response;
        }

        public Hoja4 PorcentajeDiferenciHoja4(Hoja4 diferencia, Hoja4 totalAñoAnt)
        {
            Hoja4 response = new Hoja4();
            response.Dia = "DIF %";

            try
            {
                #region % diferencia
                response.Ubre_MA = diferencia != null && totalAñoAnt != null && totalAñoAnt.Ubre_MA > 0 ? diferencia.Ubre_MA / totalAñoAnt.Ubre_MA : 0;
                response.Ubre_SL = diferencia != null && totalAñoAnt != null && totalAñoAnt.Ubre_SL > 0 ? diferencia.Ubre_SL / totalAñoAnt.Ubre_SL : 0;
                response.Metabolicos_FL = diferencia != null && totalAñoAnt != null && totalAñoAnt.Metabolicos_FL > 0 ? diferencia.Metabolicos_FL / totalAñoAnt.Metabolicos_FL : 0;
                response.Metabolicos_CET = diferencia != null && totalAñoAnt != null && totalAñoAnt.Metabolicos_CET > 0 ? diferencia.Metabolicos_CET / totalAñoAnt.Metabolicos_CET : 0;
                response.Locomotores_BE = diferencia != null && totalAñoAnt != null && totalAñoAnt.Locomotores_BE > 0 ? diferencia.Locomotores_BE / totalAñoAnt.Locomotores_BE : 0;
                response.Locomotores_TRA = diferencia != null && totalAñoAnt != null && totalAñoAnt.Locomotores_TRA > 0 ? diferencia.Locomotores_TRA / totalAñoAnt.Locomotores_TRA : 0;
                response.Locomotores_GA = diferencia != null && totalAñoAnt != null && totalAñoAnt.Locomotores_GA > 0 ? diferencia.Locomotores_GA / totalAñoAnt.Locomotores_GA : 0;
                response.Digestivos_AC = diferencia != null && totalAñoAnt != null && totalAñoAnt.Digestivos_AC > 0 ? diferencia.Digestivos_AC / totalAñoAnt.Digestivos_AC : 0;
                response.Digestivos_ES = diferencia != null && totalAñoAnt != null && totalAñoAnt.Digestivos_ES > 0 ? diferencia.Digestivos_ES / totalAñoAnt.Digestivos_ES : 0;
                response.Digestivos_DI = diferencia != null && totalAñoAnt != null && totalAñoAnt.Digestivos_DI > 0 ? diferencia.Digestivos_DI / totalAñoAnt.Digestivos_DI : 0;
                response.Digestivos_TI = diferencia != null && totalAñoAnt != null && totalAñoAnt.Digestivos_TI > 0 ? diferencia.Digestivos_TI / totalAñoAnt.Digestivos_TI : 0;
                response.Reproductivos_RE = diferencia != null && totalAñoAnt != null && totalAñoAnt.Reproductivos_RE > 0 ? diferencia.Reproductivos_RE / totalAñoAnt.Reproductivos_RE : 0;
                response.Reproductivos_ME = diferencia != null && totalAñoAnt != null && totalAñoAnt.Reproductivos_ME > 0 ? diferencia.Reproductivos_ME / totalAñoAnt.Reproductivos_ME : 0;
                response.Reproductivos_PIO = diferencia != null && totalAñoAnt != null && totalAñoAnt.Reproductivos_PIO > 0 ? diferencia.Reproductivos_PIO / totalAñoAnt.Reproductivos_PIO : 0;
                response.Reproductivos_QUI = diferencia != null && totalAñoAnt != null && totalAñoAnt.Reproductivos_QUI > 0 ? diferencia.Reproductivos_QUI / totalAñoAnt.Reproductivos_QUI : 0;
                response.Reproductivos_CS = diferencia != null && totalAñoAnt != null && totalAñoAnt.Reproductivos_CS > 0 ? diferencia.Reproductivos_CS / totalAñoAnt.Reproductivos_CS : 0;
                response.Respiratorios_Neu = diferencia != null && totalAñoAnt != null && totalAñoAnt.Respiratorios_Neu > 0 ? diferencia.Respiratorios_Neu / totalAñoAnt.Respiratorios_Neu : 0;
                response.Becerras_Neu = diferencia != null && totalAñoAnt != null && totalAñoAnt.Becerras_Neu > 0 ? diferencia.Becerras_Neu / totalAñoAnt.Becerras_Neu : 0;
                response.Becerras_Fie = diferencia != null && totalAñoAnt != null && totalAñoAnt.Becerras_Fie > 0 ? diferencia.Becerras_Fie / totalAñoAnt.Becerras_Fie : 0;
                response.Becerras_Di = diferencia != null && totalAñoAnt != null && totalAñoAnt.Becerras_Di > 0 ? diferencia.Becerras_Di / totalAñoAnt.Becerras_Di : 0;
                response.Becerras_Conj = diferencia != null && totalAñoAnt != null && totalAñoAnt.Becerras_Conj > 0 ? diferencia.Becerras_Conj / totalAñoAnt.Becerras_Conj : 0;
                response.Vacas_Diag = diferencia != null && totalAñoAnt != null && totalAñoAnt.Vacas_Diag > 0 ? diferencia.Vacas_Diag / totalAñoAnt.Vacas_Diag : 0;
                response.Vacas_Pren = diferencia != null && totalAñoAnt != null && totalAñoAnt.Vacas_Pren > 0 ? diferencia.Vacas_Pren / totalAñoAnt.Vacas_Pren : 0;
                response.Vacas_Porcentaje_Pren = diferencia != null && totalAñoAnt != null && totalAñoAnt.Vacas_Porcentaje_Pren > 0 ? diferencia.Vacas_Porcentaje_Pren / totalAñoAnt.Vacas_Porcentaje_Pren : 0;
                response.Vacas_Vacias = diferencia != null && totalAñoAnt != null && totalAñoAnt.Vacas_Vacias > 0 ? diferencia.Vacas_Vacias / totalAñoAnt.Vacas_Vacias : 0;
                response.Vacas_Porcentaje_Vacias = diferencia != null && totalAñoAnt != null && totalAñoAnt.Vacas_Porcentaje_Vacias > 0 ? diferencia.Vacas_Porcentaje_Vacias / totalAñoAnt.Vacas_Porcentaje_Vacias : 0;
                response.Vaquillas_Diag = diferencia != null && totalAñoAnt != null && totalAñoAnt.Vaquillas_Diag > 0 ? diferencia.Vaquillas_Diag / totalAñoAnt.Vaquillas_Diag : 0;
                response.Vaquillas_Pren = diferencia != null && totalAñoAnt != null && totalAñoAnt.Vaquillas_Pren > 0 ? diferencia.Vaquillas_Pren / totalAñoAnt.Vaquillas_Pren : 0;
                response.Vaquillas_Porcentaje_Pren = diferencia != null && totalAñoAnt != null && totalAñoAnt.Vaquillas_Porcentaje_Pren > 0 ? diferencia.Vaquillas_Porcentaje_Pren / totalAñoAnt.Vaquillas_Porcentaje_Pren : 0;
                response.Vaquillas_Vacias = diferencia != null && totalAñoAnt != null && totalAñoAnt.Vaquillas_Vacias > 0 ? diferencia.Vaquillas_Vacias / totalAñoAnt.Vaquillas_Vacias : 0;
                response.Vaquillas_Porcentaje_Vacias = diferencia != null && totalAñoAnt != null && totalAñoAnt.Vaquillas_Porcentaje_Vacias > 0 ? diferencia.Vaquillas_Porcentaje_Vacias / totalAñoAnt.Vaquillas_Porcentaje_Vacias : 0;
                response.Abortos_Vaquillas = diferencia != null && totalAñoAnt != null && totalAñoAnt.Abortos_Vaquillas > 0 ? diferencia.Abortos_Vaquillas / totalAñoAnt.Abortos_Vaquillas : 0;
                response.Abortos_Vacas = diferencia != null && totalAñoAnt != null && totalAñoAnt.Abortos_Vacas > 0 ? diferencia.Abortos_Vacas / totalAñoAnt.Abortos_Vacas : 0;
                #endregion
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "PorcentajeDiferenciHoja4", $"Error in {nameof(PorcentajeDiferenciHoja4)}", ex.Message, ex.Message, ex.StackTrace, guid);
            }

            return response;
        }
        #endregion

        #region espacios en blanco
        public List<T> EspaciosEnBlanco<T>(int renglones) where T : new()
        {
            List<T> response = new List<T>();
            int renglonesTotal = 31 - renglones;

            for (int i = 0; i < renglonesTotal; i++)
            {
                response.Add(new T());
            }

            return response;
        }

        #endregion

        #region renglon Total
        public Hoja1 RenglonTotalHoja1(List<Hoja1> request)
        {
            Hoja1 response = new Hoja1();

            try
            {
                response.Dia = "TOTAL";
                response.Ordeño = 0;
                response.Secas = 0;
                response.Hato = 0;
                response.Leche = 0;
                response.Antib = 0;
                response.Total = 0;

                foreach (Hoja1 item in request)
                {
                    response.Ordeño += item.Ordeño != null ? item.Ordeño : 0;
                    response.Secas += item.Secas != null ? item.Secas : 0;
                    response.Hato += item.Hato != null ? item.Hato : 0;
                    response.Leche += item.Leche != null ? item.Leche : 0;
                    response.Antib += item.Antib != null ? item.Antib : 0;
                    response.Total += item.Total != null ? item.Total : 0;
                }

            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "RenglonTotalHoja1", $"Error in {nameof(RenglonTotalHoja1)}", ex.Message, ex.Message, ex.StackTrace, guid);
            }

            return response;
        }

        public Hoja2 RenglonTotalHoja2(List<Hoja2> request)
        {
            Hoja2 response = new Hoja2();

            try
            {
                response.Dia = "TOTAL";
                response.Crianza_Inventario = 0;
                response.Jaulas_Inventario = 0;
                response.Crecimiento_Inventario = 0;
                response.Desarrollo_Inventario = 0;
                response.Secas_Inventario = 0;
                response.Reto_Inventario = 0;
                response.Inventario_Total = 0;

                foreach (Hoja2 item in request)
                {
                    response.Crianza_Inventario += item.Crianza_Inventario != null ? item.Crianza_Inventario : 0;
                    response.Jaulas_Inventario += item.Jaulas_Inventario != null ? item.Jaulas_Inventario : 0;
                    response.Crecimiento_Inventario += item.Crecimiento_Inventario != null ? item.Crecimiento_Inventario : 0;
                    response.Desarrollo_Inventario += item.Desarrollo_Inventario != null ? item.Desarrollo_Inventario : 0;
                    response.Secas_Inventario += item.Secas_Inventario != null ? item.Secas_Inventario : 0;
                    response.Reto_Inventario += item.Reto_Inventario != null ? item.Reto_Inventario : 0;
                    response.Inventario_Total += item.Inventario_Total != null ? item.Inventario_Total : 0;
                }
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "RenglonTotalHoja2", $"Error in {nameof(RenglonTotalHoja2)}", ex.Message, ex.Message, ex.StackTrace, guid);
            }

            return response;
        }
        #endregion

        #region Colorimetria hojas
        public void AsignarColorimetriaHoja1(Rancho rancho, List<Hoja1> reporte, List<CalostroYOrdeña> datosCalostro, Hoja1 promedio, DateTime fechaInicio, DateTime fechaFin, List<BacteriologiaLeche> datosBacteriologia)
        {
            try
            {
                CalostroYOrdeña busquedaCalostro = new CalostroYOrdeña();
                BacteriologiaLeche busquedaBacteriologia = new BacteriologiaLeche();
                gthEntity.IndicadorReporteTeorico busquedaIndicador = new gthEntity.IndicadorReporteTeorico();

                foreach (Hoja1 item in reporte)
                {
                    busquedaCalostro = (from x in datosCalostro where x.Fecha.Day.ToString() == item.Dia select x).ToList().FirstOrDefault();
                    busquedaBacteriologia = (from x in datosBacteriologia where x.Fecha.Day.ToString() == item.Dia select x).ToList().FirstOrDefault();
                    busquedaIndicador = (from x in indicadorProduccion where x.Fecha.Day.ToString() == item.Dia select x).ToList().FirstOrDefault();

                    busquedaCalostro = busquedaCalostro ?? new CalostroYOrdeña();
                    busquedaBacteriologia = busquedaBacteriologia ?? new BacteriologiaLeche();
                    busquedaIndicador = busquedaIndicador ?? new gthEntity.IndicadorReporteTeorico();

                    item.Color_ILCA = Color1Hoja1(item.ILCA, Convert.ToDecimal(promedioProduccion.ILCAProducccion));
                    item.Color_IC = Color1Hoja1(item.IC, Convert.ToDecimal(promedioProduccion.ICProduccion));
                    item.Color_CostoLitro = Color2Hoja1(item.Costo_Litro, promedio.Costo_Litro);

                    item.Color_MH = ColorDato(item.Ordeño, item.MH, busquedaIndicador.MateriaHumeda);
                    item.Color_PorcentajeMs = ColorDato(item.Ordeño, item.Porcentaje_MS, busquedaIndicador.PorcentajeMs / 100);
                    item.Color_MS = ColorDato(item.Ordeño, item.MS, busquedaIndicador.MateriaSeca);
                    item.Color_CostoProd = ColorDato(item.Ordeño, item.Costo_Prod, busquedaIndicador.Costo);
                    item.Color_CostoMs = ColorDato(item.Ordeño, item.Costo_MS, busquedaIndicador.PrecioKgMs);

                    item.Color_Leche = Color1Hoja1(item.Leche, promedio.Leche);
                    item.Color_Media = Color1Hoja1(item.Media, promedio.Media);
                    item.Color_Total = Color1Hoja1(item.Total, promedio.Total);
                    item.Color_Ses1 = ColorSes(rancho.Pesadores, item.Ordeño, item.NoID_Ses1);
                    item.Color_Ses2 = ColorSes(rancho.Pesadores, item.Ordeño, item.NoID_Ses2);
                    item.Color_Ses3 = ColorSes(rancho.Pesadores, item.Ordeño, item.NoID_Ses3);
                    item.Color_CTD = item.CTD != null ? busquedaBacteriologia != null ? busquedaBacteriologia.Bacteriologia < 0.2M ? "#FFC9C9" : "#F2F2F2" : string.Empty : string.Empty;
                    item.Color_ANT = ColorAntHoja1(item.Media, item.Antib, item.ANT);
                    item.Color_HoraSes1 = busquedaCalostro != null ? ColorSesionesHora(busquedaCalostro.Horario_Sesion1) : string.Empty;
                    item.Color_HoraSes2 = busquedaCalostro != null ? ColorSesionesHora(busquedaCalostro.Horario_Sesion2) : string.Empty;
                    item.Color_HoraSes3 = busquedaCalostro != null ? ColorSesionesHora(busquedaCalostro.Horario_Sesion3) : string.Empty;
                    item.Color_Ordeño = busquedaCalostro != null ? ColorSalaOrdeño(busquedaCalostro.Capacidad_Ordeño) : string.Empty;

                    if (item.Color_HoraSes1 != "#FFC9C9" && item.Color_HoraSes1 != string.Empty && item.Color_HoraSes1 != "")
                        item.Color_HoraSes1 = busquedaCalostro != null ? busquedaCalostro.Horario_Dif_Sesion1 < 10 ? "#FFC9C9" : item.Color_HoraSes1 : item.Color_HoraSes1;

                    if (item.Color_HoraSes2 != "#FFC9C9" && item.Color_HoraSes1 != string.Empty && item.Color_HoraSes1 != "")
                        item.Color_HoraSes2 = busquedaCalostro != null ? busquedaCalostro.Horario_Dif_Sesion2 < 10 ? "#FFC9C9" : item.Color_HoraSes2 : item.Color_HoraSes2;

                    if (item.Color_HoraSes3 != "#FFC9C9" && item.Color_HoraSes1 != string.Empty && item.Color_HoraSes1 != "")
                        item.Color_HoraSes3 = busquedaCalostro != null ? busquedaCalostro.Horario_Dif_Sesion3 < 10 ? "#FFC9C9" : item.Color_HoraSes3 : item.Color_HoraSes3;
                }
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "ColorimetriaHoja1", $"Error in {nameof(AsignarColorimetriaHoja1)}", ex.Message, ex.Message, ex.StackTrace, guid);
            }
        }

        public void AsignarColorimetriaHoja2(List<Hoja2> reporte, Utilidad utilidad)
        {
            gthEntity.IndicadorReporteTeorico busquedaindicadorCrecimiento = new gthEntity.IndicadorReporteTeorico();
            gthEntity.IndicadorReporteTeorico busquedaindicadorDesarrollo = new gthEntity.IndicadorReporteTeorico();
            gthEntity.IndicadorReporteTeorico busquedaindicadorVaquillas = new gthEntity.IndicadorReporteTeorico();
            gthEntity.IndicadorReporteTeorico busquedaindicadorSecas = new gthEntity.IndicadorReporteTeorico();
            gthEntity.IndicadorReporteTeorico busquedaindicadorReto = new gthEntity.IndicadorReporteTeorico();

            try
            {
                foreach (Hoja2 item in reporte)
                {
                    busquedaindicadorCrecimiento = (from x in indicadorCrecimiento where x.Fecha.Day.ToString() == item.Dia select x).ToList().FirstOrDefault();
                    busquedaindicadorDesarrollo = (from x in indicadorDesarrollo where x.Fecha.Day.ToString() == item.Dia select x).ToList().FirstOrDefault();
                    busquedaindicadorVaquillas = (from x in indicadorVaquillas where x.Fecha.Day.ToString() == item.Dia select x).ToList().FirstOrDefault();
                    busquedaindicadorSecas = (from x in indicadorSecas where x.Fecha.Day.ToString() == item.Dia select x).ToList().FirstOrDefault();
                    busquedaindicadorReto = (from x in indicadorReto where x.Fecha.Day.ToString() == item.Dia select x).ToList().FirstOrDefault();

                    busquedaindicadorCrecimiento = busquedaindicadorCrecimiento ?? new gthEntity.IndicadorReporteTeorico();
                    busquedaindicadorDesarrollo = busquedaindicadorDesarrollo ?? new gthEntity.IndicadorReporteTeorico();
                    busquedaindicadorVaquillas = busquedaindicadorVaquillas ?? new gthEntity.IndicadorReporteTeorico();
                    busquedaindicadorSecas = busquedaindicadorSecas ?? new gthEntity.IndicadorReporteTeorico();
                    busquedaindicadorReto = busquedaindicadorReto ?? new gthEntity.IndicadorReporteTeorico();

                    item.Color_MH_Crecimiento = busquedaindicadorCrecimiento != null ? ColorDato(item.Crecimiento_Inventario, item.Crecimiento_MH, busquedaindicadorCrecimiento.MateriaHumeda) : "";
                    item.Color_MS_Crecimiento = busquedaindicadorCrecimiento != null ? ColorDato(item.Crecimiento_Inventario, item.Crecimiento_MS, busquedaindicadorCrecimiento.MateriaSeca) : "";
                    item.Color_PorcentajeMs_Crecimiento = busquedaindicadorCrecimiento != null ? ColorDato(item.Crecimiento_Inventario, item.Crecimiento_PorcentajeMS, busquedaindicadorCrecimiento.PorcentajeMs) : "";
                    item.Color_CostoMS_Crecimiento = busquedaindicadorCrecimiento != null ? ColorDato(item.Crecimiento_Inventario, item.Crecimiento_CostoMS, busquedaindicadorCrecimiento.PrecioKgMs) : "";
                    item.Color_Costo_Crecimiento = busquedaindicadorCrecimiento != null ? ColorDato(item.Crecimiento_Inventario, item.Crecimiento_Costo, busquedaindicadorCrecimiento.Costo) : "";

                    item.Color_MH_Desarrollo = busquedaindicadorDesarrollo != null ? ColorDato(item.Desarrollo_Inventario, item.Desarrollo_MH, busquedaindicadorDesarrollo.MateriaHumeda) : "";
                    item.Color_MS_Desarrollo = busquedaindicadorDesarrollo != null ? ColorDato(item.Desarrollo_Inventario, item.Desarrollo_MS, busquedaindicadorDesarrollo.MateriaSeca) : "";
                    item.Color_PorcentajeMs_Desarrollo = busquedaindicadorDesarrollo != null ? ColorDato(item.Desarrollo_Inventario, item.Desarrollo_PorcentajeMS, busquedaindicadorDesarrollo.PorcentajeMs) : "";
                    item.Color_CostoMS_Desarrollo = busquedaindicadorDesarrollo != null ? ColorDato(item.Desarrollo_Inventario, item.Desarrollo_CostoMS, busquedaindicadorDesarrollo.PrecioKgMs) : "";
                    item.Color_Costo_Desarrollo = busquedaindicadorDesarrollo != null ? ColorDato(item.Desarrollo_Inventario, item.Desarrollo_Costo, busquedaindicadorDesarrollo.Costo) : "";

                    item.Color_MH_Vaquillas = busquedaindicadorVaquillas != null ? ColorDato(item.Vaquillas_Inventario, item.Vaquillas_MH, busquedaindicadorVaquillas.MateriaHumeda) : "";
                    item.Color_MS_Vaquillas = busquedaindicadorVaquillas != null ? ColorDato(item.Vaquillas_Inventario, item.Vaquillas_MS, busquedaindicadorVaquillas.MateriaSeca) : "";
                    item.Color_PorcentajeMs_Vaquillas = busquedaindicadorVaquillas != null ? ColorDato(item.Vaquillas_Inventario, item.Vaquillas_PorcentajeMS, busquedaindicadorVaquillas.PorcentajeMs) : "";
                    item.Color_CostoMS_Vaquillas = busquedaindicadorVaquillas != null ? ColorDato(item.Vaquillas_Inventario, item.Vaquillas_CostoMS, busquedaindicadorVaquillas.PrecioKgMs) : "";
                    item.Color_Costo_Vaquillas = busquedaindicadorVaquillas != null ? ColorDato(item.Vaquillas_Inventario, item.Vaquillas_Costo, busquedaindicadorVaquillas.Costo) : "";

                    item.Color_MH_Secas = busquedaindicadorSecas != null ? ColorDato(item.Secas_Inventario, item.Secas_MH, busquedaindicadorSecas.MateriaHumeda) : "";
                    item.Color_MS_Secas = busquedaindicadorSecas != null ? ColorDato(item.Secas_Inventario, item.Secas_MS, busquedaindicadorSecas.MateriaSeca) : "";
                    item.Color_PorcentajeMs_Secas = busquedaindicadorSecas != null ? ColorDato(item.Secas_Inventario, item.Secas_PorcentajeMS, busquedaindicadorSecas.PorcentajeMs) : "";
                    item.Color_CostoMS_Secas = busquedaindicadorSecas != null ? ColorDato(item.Secas_Inventario, item.Secas_CostoMS, busquedaindicadorSecas.PrecioKgMs) : "";
                    item.Color_Costo_Secas = busquedaindicadorSecas != null ? ColorDato(item.Secas_Inventario, item.Secas_Costo, busquedaindicadorSecas.Costo) : "";

                    item.Color_MH_Reto = busquedaindicadorReto != null ? ColorDato(item.Reto_Inventario, item.Reto_MH, busquedaindicadorReto.MateriaHumeda) : "";
                    item.Color_MS_Reto = busquedaindicadorReto != null ? ColorDato(item.Reto_Inventario, item.Reto_MS, busquedaindicadorReto.MateriaSeca) : "";
                    item.Color_PorcentajeMs_Reto = busquedaindicadorReto != null ? ColorDato(item.Reto_Inventario, item.Reto_PorcentajeMS, busquedaindicadorReto.PorcentajeMs) : "";
                    item.Color_CostoMS_Reto = busquedaindicadorReto != null ? ColorDato(item.Reto_Inventario, item.Reto_CostoMS, busquedaindicadorReto.PrecioKgMs) : "";
                    item.Color_Costo_Reto = busquedaindicadorReto != null ? ColorDato(item.Reto_Inventario, item.Reto_Costo, busquedaindicadorReto.Costo) : "";

                    item.Color_IXA = ColorUtilidad(item.IngresoxAnimal, utilidad.IXA, false);
                    item.Color_CXA = ColorUtilidad(item.CostoxAnimal, utilidad.CXA, true);
                    item.Color_PorcentajeCXA = ColorUtilidad(item.Porcentaje_CostoxAnimal, utilidad.PORCENTAJE_C, true);
                    item.Color_UXA = ColorUtilidad(item.UtilidadxAnimal, utilidad.UXA, false);
                    item.Color_PorcentajeUXA = ColorUtilidad(item.Porcentaje_UtilidadxAnimal, utilidad.PORCENTAJE_U, false);
                }
            }
            catch { }

        }
        #endregion

        #region metodos para asignar colores
        private string Color1Hoja1(decimal? valor, decimal? promedio)
        {
            string color = "";

            if (valor != 0)
            {
                if (!(promedio == 0 || promedio == null))
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

        private string Color2Hoja1(decimal? valor, decimal? promedio)
        {
            string color = "";

            if (valor != 0)
            {
                if (!(promedio == 0 || promedio == null))
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
            }
            return color;
        }

        private string ColorAntHoja1(decimal? media, decimal? antib, decimal? ant)
        {
            string color = "";

            try
            {
                decimal? result = ant > 0 ? antib / ant : 0;

                switch (result)
                {
                    //Color Verde
                    case decimal n when (n >= (media * 0.5M) && n <= (media * 0.7M)):
                        color = "#DEEDD3";
                        break;
                    //Color Amarillo
                    case decimal n when ((n > (media * 0.7M) && n <= (media * 0.8M)) || (n >= (media * 0.4M) && n < (media * 0.5M))):
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

        private string ColorSes(bool pesadores, decimal? inventario, decimal? vacasSesion)
        {
            string color = "";

            try
            {
                if (pesadores)
                {
                    if (inventario > 0)
                    {
                        decimal? porcentaje = vacasSesion / inventario * 100;

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
            }
            catch { }

            return color;
        }

        private string ColorSesionesHora(decimal? valor)
        {
            string color = "";

            if (valor != null)
            {
                switch (valor)
                {
                    //color rojo
                    case decimal n when (n < -5 || n > 15):
                        color = "#FFC9C9";
                        break;
                    //color verde
                    case decimal n when (n >= -5 && n <= 10):
                        color = "#DEEDD3";
                        break;
                    //color amarillo 
                    case decimal n when (n > 10 || n <= 15):
                        color = "#FFF5D9";
                        break;
                }
            }

            return color;
        }

        private string ColorDato(decimal? inventario, decimal? valor, decimal? promedio)
        {
            string color = "";

            try
            {
                if (inventario != null)
                {
                    if (inventario != 0)
                    {
                        if (valor != null && valor != 0)
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
                        else
                            color = "#FFC9C9";

                    }
                }
            }
            catch { }

            return color;
        }

        private string ColorSalaOrdeño(decimal? promedio)
        {
            string color = "";

            if (!(promedio == 0 && promedio == null))
            {
                switch (promedio)
                {
                    //color verde
                    case decimal n when n <= 94:
                        color = "#DEEDD3";
                        break;
                    //color amarillo 
                    case decimal n when (n > 94 && n <= 98):
                        color = "#FFF5D9";
                        break;
                    //color rojo
                    case decimal n when n > 98:
                        color = "#FFC9C9";
                        break;
                }
            }


            return color;
        }

        private string ColorUtilidad(decimal? valor, decimal? promedio, bool esCosto)
        {
            string color = "";

            try
            {
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
            }
            catch { }

            return color;
        }
        #endregion

        #region metodos para quitar ceros
        public void QuitarCeros(List<Hoja1> request)
        {
            foreach (Hoja1 item in request)
            {
                item.Ordeño = item.Ordeño == 0 ? null : item.Ordeño;

                item.Secas = item.Secas == 0 ? null : item.Secas;
                item.Hato = item.Hato == 0 ? null : item.Hato;
                item.Porcentaje_Lact = item.Porcentaje_Lact == 0 ? null : item.Porcentaje_Lact;
                item.Porcentaje_Prot = item.Porcentaje_Prot == 0 ? null : item.Porcentaje_Prot;
                item.Urea = item.Urea == 0 ? null : item.Urea;
                item.Porcentaje_Grasa = item.Porcentaje_Grasa == 0 ? null : item.Porcentaje_Grasa;
                item.CCS = item.CCS == 0 ? null : item.CCS;
                item.CTD = item.CTD == 0 ? null : item.CTD;
                item.Color_CTD = item.CTD == null ? "" : item.Color_CTD;

                item.Leche = item.Leche == 0 ? null : item.Leche;
                item.Antib = item.Antib == 0 ? null : item.Antib;
                item.Media = item.Media == 0 ? null : item.Media;
                item.Total = item.Total == 0 ? null : item.Total;
                item.DEL = item.DEL == 0 ? null : item.DEL;
                item.ANT = item.ANT == 0 ? null : item.ANT;

                item.EA = item.EA == 0 ? null : item.EA;
                item.ILCA = item.ILCA == 0 ? null : item.ILCA;
                item.IC = item.IC == 0 ? null : item.IC;
                item.Costo_Litro = item.Costo_Litro == 0 ? null : item.Costo_Litro;
                item.MH = item.MH == 0 ? null : item.MH;
                item.MS = item.MS == 0 ? null : item.MS;
                item.Porcentaje_MS = item.Porcentaje_MS == 0 ? null : item.Porcentaje_MS;
                item.SA = item.SA == 0 ? null : item.SA;
                item.MSS = item.MSS == 0 ? null : item.MSS;
                item.EAS = item.EAS == 0 ? null : item.EAS;
                item.Porcentaje_Sob = item.Porcentaje_Sob == 0 ? null : item.Porcentaje_Sob;
                item.Costo_Prod = item.Costo_Prod == 0 ? null : item.Costo_Prod;
                item.Costo_MS = item.Costo_MS == 0 ? null : item.Costo_MS;

                item.Cribas_N1 = item.Cribas_N1 == 0 ? null : item.Cribas_N1;
                item.Cribas_N2 = item.Cribas_N2 == 0 ? null : item.Cribas_N2;
                item.Cribas_N3 = item.Cribas_N3 == 0 ? null : item.Cribas_N3;
                item.Cribas_N4 = item.Cribas_N4 == 0 ? null : item.Cribas_N4;

                item.NoID_Ses1 = item.NoID_Ses1 == 0 ? null : item.NoID_Ses1;
                item.NoID_Ses2 = item.NoID_Ses2 == 0 ? null : item.NoID_Ses2;
                item.NoID_Ses3 = item.NoID_Ses3 == 0 ? null : item.NoID_Ses3;

                item.SES1 = item.SES1 == 0 ? null : item.SES1;
                item.SES2 = item.SES2 == 0 ? null : item.SES2;
                item.SES3 = item.SES3 == 0 ? null : item.SES3;

                item.Porcentaje_Revueltas = item.Porcentaje_Revueltas == 0 ? null : item.Porcentaje_Revueltas;

                item.Color_Ses1 = item.NoID_Ses1 == null ? "" : item.Color_Ses1;
                item.Color_Ses2 = item.NoID_Ses2 == null ? "" : item.Color_Ses2;
                item.Color_Ses3 = item.NoID_Ses3 == null ? "" : item.Color_Ses3;

                item.Color_HoraSes1 = item.SES1 == null ? "" : item.Color_HoraSes1;
                item.Color_HoraSes2 = item.SES2 == null ? "" : item.Color_HoraSes2;
                item.Color_HoraSes3 = item.SES3 == null ? "" : item.Color_HoraSes3;
                item.Color_Ordeño = item.Ordeño == null ? "" : item.Color_Ordeño;
                item.Color_ANT = item.ANT == null ? "" : item.Color_ANT;
            }
        }

        public void QuitarCeros(Hoja1 request)
        {
            try
            {
                request.Ordeño = request.Ordeño == 0 ? null : request.Ordeño;
                request.Secas = request.Secas == 0 ? null : request.Secas;
                request.Hato = request.Hato == 0 ? null : request.Hato;
                request.Porcentaje_Lact = request.Porcentaje_Lact == 0 ? null : request.Porcentaje_Lact;
                request.Porcentaje_Prot = request.Porcentaje_Prot == 0 ? null : request.Porcentaje_Prot;
                request.Urea = request.Urea == 0 ? null : request.Urea;
                request.Porcentaje_Grasa = request.Porcentaje_Grasa == 0 ? null : request.Porcentaje_Grasa;
                request.CCS = request.CCS == 0 ? null : request.CCS;
                request.CTD = request.CTD == 0 ? null : request.CTD;

                request.Leche = request.Leche == 0 ? null : request.Leche;
                request.Antib = request.Antib == 0 ? null : request.Antib;
                request.Media = request.Media == 0 ? null : request.Media;
                request.Total = request.Total == 0 ? null : request.Total;
                request.DEL = request.DEL == 0 ? null : request.DEL;
                request.ANT = request.ANT == 0 ? null : request.ANT;

                request.EA = request.EA == 0 ? null : request.EA;
                request.ILCA = request.ILCA == 0 ? null : request.ILCA;
                request.IC = request.IC == 0 ? null : request.IC;
                request.Costo_Litro = request.Costo_Litro == 0 ? null : request.Costo_Litro;
                request.MH = request.MH == 0 ? null : request.MH;
                request.MS = request.MS == 0 ? null : request.MS;
                request.Porcentaje_MS = request.Porcentaje_MS == 0 ? null : request.Porcentaje_MS;
                request.SA = request.SA == 0 ? null : request.SA;
                request.MSS = request.MSS == 0 ? null : request.MSS;
                request.EAS = request.EAS == 0 ? null : request.EAS;
                request.Porcentaje_Sob = request.Porcentaje_Sob == 0 ? null : request.Porcentaje_Sob;
                request.Costo_Prod = request.Costo_Prod == 0 ? null : request.Costo_Prod;
                request.Costo_MS = request.Costo_MS == 0 ? null : request.Costo_MS;

                request.Cribas_N1 = request.Cribas_N1 == 0 ? null : request.Cribas_N1;
                request.Cribas_N2 = request.Cribas_N2 == 0 ? null : request.Cribas_N2;
                request.Cribas_N3 = request.Cribas_N3 == 0 ? null : request.Cribas_N3;
                request.Cribas_N4 = request.Cribas_N4 == 0 ? null : request.Cribas_N4;

                request.NoID_Ses1 = request.NoID_Ses1 == 0 ? null : request.NoID_Ses1;
                request.NoID_Ses2 = request.NoID_Ses2 == 0 ? null : request.NoID_Ses2;
                request.NoID_Ses3 = request.NoID_Ses3 == 0 ? null : request.NoID_Ses3;

                request.SES1 = request.SES1 == 0 ? null : request.SES1;
                request.SES2 = request.SES2 == 0 ? null : request.SES2;
                request.SES3 = request.SES3 == 0 ? null : request.SES3;

                request.Porcentaje_Revueltas = request.Porcentaje_Revueltas == 0 ? null : request.Porcentaje_Revueltas;
            }
            catch { }
        }

        public void QuitarCeros(List<Hoja2> reporte)
        {
            foreach (Hoja2 item in reporte)
            {
                #region Crianza
                item.Crianza_Inventario = item.Crianza_Inventario == 0 ? null : item.Crianza_Inventario;
                #endregion

                #region Jaulas 
                item.Jaulas_Inventario = item.Jaulas_Inventario == 0 ? null : item.Jaulas_Inventario;
                item.Jaulas_Costo = item.Jaulas_Costo == 0 ? null : item.Jaulas_Costo;
                #endregion

                #region Crecimiento 
                item.Crecimiento_Inventario = item.Crecimiento_Inventario == 0 ? null : item.Crecimiento_Inventario;
                item.Crecimiento_MH = item.Crecimiento_MH == 0 ? null : item.Crecimiento_MH;
                item.Crecimiento_Costo = item.Crecimiento_Costo == 0 ? null : item.Crecimiento_Costo;
                item.Crecimiento_PorcentajeMS = item.Crecimiento_PorcentajeMS == 0 ? null : item.Crecimiento_PorcentajeMS;
                item.Crecimiento_MS = item.Crecimiento_MS == 0 ? null : item.Crecimiento_MS;
                item.Crecimiento_CostoMS = item.Crecimiento_CostoMS == 0 ? null : item.Crecimiento_CostoMS;
                #endregion

                #region Desarrollo 
                item.Desarrollo_Inventario = item.Desarrollo_Inventario == 0 ? null : item.Desarrollo_Inventario;
                item.Desarrollo_MH = item.Desarrollo_MH == 0 ? null : item.Desarrollo_MH;
                item.Desarrollo_Costo = item.Desarrollo_Costo == 0 ? null : item.Desarrollo_Costo;
                item.Desarrollo_PorcentajeMS = item.Desarrollo_PorcentajeMS == 0 ? null : item.Desarrollo_PorcentajeMS;
                item.Desarrollo_MS = item.Desarrollo_MS == 0 ? null : item.Desarrollo_MS;
                item.Desarrollo_CostoMS = item.Desarrollo_CostoMS == 0 ? null : item.Desarrollo_CostoMS;
                #endregion

                #region Vaquillas 
                item.Vaquillas_Inventario = item.Vaquillas_Inventario == 0 ? null : item.Vaquillas_Inventario;
                item.Vaquillas_MH = item.Vaquillas_MH == 0 ? null : item.Vaquillas_MH;
                item.Vaquillas_Costo = item.Vaquillas_Costo == 0 ? null : item.Vaquillas_Costo;
                item.Vaquillas_PorcentajeMS = item.Vaquillas_PorcentajeMS == 0 ? null : item.Vaquillas_PorcentajeMS;
                item.Vaquillas_MS = item.Vaquillas_MS == 0 ? null : item.Vaquillas_MS;
                item.Vaquillas_CostoMS = item.Vaquillas_CostoMS == 0 ? null : item.Vaquillas_CostoMS;
                #endregion

                #region Secas 
                item.Secas_Inventario = item.Secas_Inventario == 0 ? null : item.Secas_Inventario;
                item.Secas_MH = item.Secas_MH == 0 ? null : item.Secas_MH;
                item.Secas_PorcentajeMS = item.Secas_PorcentajeMS == 0 ? null : item.Secas_PorcentajeMS;
                item.Secas_MS = item.Secas_MS == 0 ? null : item.Secas_MS;
                item.Secas_SA = item.Secas_SA == 0 ? null : item.Secas_SA;
                item.Secas_Mss = item.Secas_Mss == 0 ? null : item.Secas_Mss;
                item.Secas_PorcentajeSob = item.Secas_PorcentajeSob == 0 ? null : item.Secas_PorcentajeSob;
                item.Secas_Costo = item.Secas_Costo == 0 ? null : item.Secas_Costo;
                item.Secas_CostoMS = item.Secas_CostoMS == 0 ? null : item.Secas_CostoMS;
                #endregion

                #region Reto 
                item.Reto_Inventario = item.Reto_Inventario == 0 ? null : item.Reto_Inventario;
                item.Reto_MH = item.Reto_MH == 0 ? null : item.Reto_MH;
                item.Reto_PorcentajeMS = item.Reto_PorcentajeMS == 0 ? null : item.Reto_PorcentajeMS;
                item.Reto_MS = item.Reto_MS == 0 ? null : item.Reto_MS;
                item.Reto_SA = item.Reto_SA == 0 ? null : item.Reto_SA;
                item.Reto_Mss = item.Reto_Mss == 0 ? null : item.Reto_Mss;
                item.Reto_PorcentajeSob = item.Reto_PorcentajeSob == 0 ? null : item.Reto_PorcentajeSob;
                item.Reto_Costo = item.Reto_Costo == 0 ? null : item.Reto_Costo;
                item.Reto_CostoMS = item.Reto_CostoMS == 0 ? null : item.Reto_CostoMS;
                #endregion

                #region Utilidad 
                item.Inventario_Total = item.Inventario_Total == 0 ? null : item.Inventario_Total;
                item.IngresoxAnimal = item.IngresoxAnimal == 0 ? null : item.IngresoxAnimal;
                item.CostoxAnimal = item.CostoxAnimal == 0 ? null : item.CostoxAnimal;
                item.UtilidadxAnimal = item.UtilidadxAnimal == 0 ? null : item.UtilidadxAnimal;
                item.Porcentaje_CostoxAnimal = item.Porcentaje_CostoxAnimal == 0 ? null : item.Porcentaje_CostoxAnimal;
                item.Porcentaje_UtilidadxAnimal = item.Porcentaje_UtilidadxAnimal == 0 ? null : item.Porcentaje_UtilidadxAnimal;

                #endregion

            }
        }

        public void QuitarCeros(Hoja2 item)
        {
            try
            {
                #region Crianza 
                item.Crianza_Inventario = item.Crianza_Inventario == 0 ? null : item.Crianza_Inventario;
                #endregion

                #region Jaulas 
                item.Jaulas_Inventario = item.Jaulas_Inventario == 0 ? null : item.Jaulas_Inventario;
                item.Jaulas_Costo = item.Jaulas_Costo == 0 ? null : item.Jaulas_Costo;
                #endregion

                #region Crecimiento 
                item.Crecimiento_Inventario = item.Crecimiento_Inventario == 0 ? null : item.Crecimiento_Inventario;
                item.Crecimiento_MH = item.Crecimiento_MH == 0 ? null : item.Crecimiento_MH;
                item.Crecimiento_Costo = item.Crecimiento_Costo == 0 ? null : item.Crecimiento_Costo;
                item.Crecimiento_PorcentajeMS = item.Crecimiento_PorcentajeMS == 0 ? null : item.Crecimiento_PorcentajeMS;
                item.Crecimiento_MS = item.Crecimiento_MS == 0 ? null : item.Crecimiento_MS;
                item.Crecimiento_CostoMS = item.Crecimiento_CostoMS == 0 ? null : item.Crecimiento_CostoMS;
                #endregion

                #region Desarrollo 
                item.Desarrollo_Inventario = item.Desarrollo_Inventario == 0 ? null : item.Desarrollo_Inventario;
                item.Desarrollo_MH = item.Desarrollo_MH == 0 ? null : item.Desarrollo_MH;
                item.Desarrollo_Costo = item.Desarrollo_Costo == 0 ? null : item.Desarrollo_Costo;
                item.Desarrollo_PorcentajeMS = item.Desarrollo_PorcentajeMS == 0 ? null : item.Desarrollo_PorcentajeMS;
                item.Desarrollo_MS = item.Desarrollo_MS == 0 ? null : item.Desarrollo_MS;
                item.Desarrollo_CostoMS = item.Desarrollo_CostoMS == 0 ? null : item.Desarrollo_CostoMS;
                #endregion

                #region Vaquillas 
                item.Vaquillas_Inventario = item.Vaquillas_Inventario == 0 ? null : item.Vaquillas_Inventario;
                item.Vaquillas_MH = item.Vaquillas_MH == 0 ? null : item.Vaquillas_MH;
                item.Vaquillas_Costo = item.Vaquillas_Costo == 0 ? null : item.Vaquillas_Costo;
                item.Vaquillas_PorcentajeMS = item.Vaquillas_PorcentajeMS == 0 ? null : item.Vaquillas_PorcentajeMS;
                item.Vaquillas_MS = item.Vaquillas_MS == 0 ? null : item.Vaquillas_MS;
                item.Vaquillas_CostoMS = item.Vaquillas_CostoMS == 0 ? null : item.Vaquillas_CostoMS;
                #endregion

                #region Secas 
                item.Secas_Inventario = item.Secas_Inventario == 0 ? null : item.Secas_Inventario;
                item.Secas_MH = item.Secas_MH == 0 ? null : item.Secas_MH;
                item.Secas_PorcentajeMS = item.Secas_PorcentajeMS == 0 ? null : item.Secas_PorcentajeMS;
                item.Secas_MS = item.Secas_MS == 0 ? null : item.Secas_MS;
                item.Secas_SA = item.Secas_SA == 0 ? null : item.Secas_SA;
                item.Secas_Mss = item.Secas_Mss == 0 ? null : item.Secas_Mss;
                item.Secas_PorcentajeSob = item.Secas_PorcentajeSob == 0 ? null : item.Secas_PorcentajeSob;
                item.Secas_Costo = item.Secas_Costo == 0 ? null : item.Secas_Costo;
                item.Secas_CostoMS = item.Secas_CostoMS == 0 ? null : item.Secas_CostoMS;
                #endregion

                #region Reto 
                item.Reto_Inventario = item.Reto_Inventario == 0 ? null : item.Reto_Inventario;
                item.Reto_MH = item.Reto_MH == 0 ? null : item.Reto_MH;
                item.Reto_PorcentajeMS = item.Reto_PorcentajeMS == 0 ? null : item.Reto_PorcentajeMS;
                item.Reto_MS = item.Reto_MS == 0 ? null : item.Reto_MS;
                item.Reto_SA = item.Reto_SA == 0 ? null : item.Reto_SA;
                item.Reto_Mss = item.Reto_Mss == 0 ? null : item.Reto_Mss;
                item.Reto_PorcentajeSob = item.Reto_PorcentajeSob == 0 ? null : item.Reto_PorcentajeSob;
                item.Reto_Costo = item.Reto_Costo == 0 ? null : item.Reto_Costo;
                item.Reto_CostoMS = item.Reto_CostoMS == 0 ? null : item.Reto_CostoMS;
                #endregion

                #region Utilidad 
                item.Inventario_Total = item.Inventario_Total == 0 ? null : item.Inventario_Total;
                item.IngresoxAnimal = item.IngresoxAnimal == 0 ? null : item.IngresoxAnimal;
                item.CostoxAnimal = item.CostoxAnimal == 0 ? null : item.CostoxAnimal;
                item.UtilidadxAnimal = item.UtilidadxAnimal == 0 ? null : item.UtilidadxAnimal;
                item.Porcentaje_CostoxAnimal = item.Porcentaje_CostoxAnimal == 0 ? null : item.Porcentaje_CostoxAnimal;
                item.Porcentaje_UtilidadxAnimal = item.Porcentaje_UtilidadxAnimal == 0 ? null : item.Porcentaje_UtilidadxAnimal;

                #endregion

            }
            catch { }

        }

        public void QuitarCeros(List<Hoja3> reporte)
        {
            foreach (Hoja3 item in reporte)
            {
                try
                {
                    item.Bp_Gem = item.Bp_Gem == 0 ? null : item.Bp_Gem;
                    item.Jaulas_Vivas = item.Jaulas_Vivas == 0 ? null : item.Jaulas_Vivas;
                    item.Jaulas_Muertas = item.Jaulas_Muertas == 0 ? null : item.Jaulas_Muertas;

                    item.Destete_Vivas = item.Destete_Vivas == 0 ? null : item.Destete_Vivas;
                    item.Destete_Muertas = item.Destete_Muertas == 0 ? null : item.Destete_Muertas;

                    item.Vaquillas_Muertas = item.Vaquillas_Muertas == 0 ? null : item.Vaquillas_Muertas;
                    item.Vaquillas_Urgente = item.Vaquillas_Urgente == 0 ? null : item.Vaquillas_Urgente;
                    item.Vaquillas_RF = item.Vaquillas_RF == 0 ? null : item.Vaquillas_RF;
                    item.Vaquillas_RR = item.Vaquillas_RR == 0 ? null : item.Vaquillas_RR;
                    item.Vaquillas_RG = item.Vaquillas_RG == 0 ? null : item.Vaquillas_RG;
                    item.Vaquillas_Otros = item.Vaquillas_Otros == 0 ? null : item.Vaquillas_Otros;

                    item.Vacas_Muertas = item.Vacas_Muertas == 0 ? null : item.Vacas_Muertas;
                    item.Vacas_Urgente = item.Vacas_Urgente == 0 ? null : item.Vacas_Urgente;
                    item.Vacas_RF = item.Vacas_RF == 0 ? null : item.Vacas_RF;
                    item.Vacas_RR = item.Vacas_RR == 0 ? null : item.Vacas_RR;
                    item.Vacas_RG = item.Vacas_RG == 0 ? null : item.Vacas_RG;
                    item.Vacas_Otros = item.Vacas_Otros == 0 ? null : item.Vacas_Otros;

                    item.Vaquillas_ND = item.Vaquillas_ND == 0 ? null : item.Vaquillas_ND;
                    item.Vaquillas_A = item.Vaquillas_A == 0 ? null : item.Vaquillas_A;
                    item.Vaquillas_Vivos_H = item.Vaquillas_Vivos_H == 0 ? null : item.Vaquillas_Vivos_H;
                    item.Vaquillas_Vivos_M = item.Vaquillas_Vivos_M == 0 ? null : item.Vaquillas_Vivos_M;

                    item.Vacas_ND = item.Vacas_ND == 0 ? null : item.Vacas_ND;
                    item.Vacas_A = item.Vacas_A == 0 ? null : item.Vacas_A;
                    item.Vacas_Vivos_H = item.Vacas_Vivos_H == 0 ? null : item.Vacas_Vivos_H;
                    item.Vacas_Vivos_M = item.Vacas_Vivos_M == 0 ? null : item.Vacas_Vivos_M;

                    item.SC = item.SC == 0 ? null : item.SC;

                    item.Partos_Vaquillas = item.Partos_Vaquillas == 0 ? null : item.Partos_Vaquillas;
                    item.Partos_Vacas = item.Partos_Vacas == 0 ? null : item.Partos_Vacas;

                    item.Muertas_Dia = item.Muertas_Dia == 0 ? null : item.Muertas_Dia;
                    item.Muertas_Noc = item.Muertas_Noc == 0 ? null : item.Muertas_Noc;

                    item.Diferencia_Calostro = item.Diferencia_Calostro == 0 ? null : item.Diferencia_Calostro;
                    item.Porcentaje_Calostro = item.Porcentaje_Calostro == 0 ? null : item.Porcentaje_Calostro;
                    item.Calidad_Calostro = item.Calidad_Calostro == 0 ? null : item.Calidad_Calostro;

                    if (item.Diferencia_Calostro == 0)
                    {
                        item.Diferencia_Calostro = null;

                    }

                }
                catch { }
            }
        }

        public void QuitarCeros(Hoja3 item)
        {
            try
            {
                item.Jaulas_Vivas = item.Jaulas_Vivas == 0 ? null : item.Jaulas_Vivas;
                item.Jaulas_Muertas = item.Jaulas_Muertas == 0 ? null : item.Jaulas_Muertas;

                item.Destete_Vivas = item.Destete_Vivas == 0 ? null : item.Destete_Vivas;
                item.Destete_Muertas = item.Destete_Muertas == 0 ? null : item.Destete_Muertas;

                item.Vaquillas_Muertas = item.Vaquillas_Muertas == 0 ? null : item.Vaquillas_Muertas;
                item.Vaquillas_Urgente = item.Vaquillas_Urgente == 0 ? null : item.Vaquillas_Urgente;
                item.Vaquillas_RF = item.Vaquillas_RF == 0 ? null : item.Vaquillas_RF;
                item.Vaquillas_RR = item.Vaquillas_RR == 0 ? null : item.Vaquillas_RR;
                item.Vaquillas_RG = item.Vaquillas_RG == 0 ? null : item.Vaquillas_RG;
                item.Vaquillas_Otros = item.Vaquillas_Otros == 0 ? null : item.Vaquillas_Otros;

                item.Vacas_Muertas = item.Vacas_Muertas == 0 ? null : item.Vacas_Muertas;
                item.Vacas_Urgente = item.Vacas_Urgente == 0 ? null : item.Vacas_Urgente;
                item.Vacas_RF = item.Vacas_RF == 0 ? null : item.Vacas_RF;
                item.Vacas_RR = item.Vacas_RR == 0 ? null : item.Vacas_RR;
                item.Vacas_RG = item.Vacas_RG == 0 ? null : item.Vacas_RG;
                item.Vacas_Otros = item.Vacas_Otros == 0 ? null : item.Vacas_Otros;

                item.Vaquillas_ND = item.Vaquillas_ND == 0 ? null : item.Vaquillas_ND;
                item.Vaquillas_A = item.Vaquillas_A == 0 ? null : item.Vaquillas_A;
                item.Vaquillas_Vivos_H = item.Vaquillas_Vivos_H == 0 ? null : item.Vaquillas_Vivos_H;
                item.Vaquillas_Vivos_M = item.Vaquillas_Vivos_M == 0 ? null : item.Vaquillas_Vivos_M;

                item.Vacas_ND = item.Vacas_ND == 0 ? null : item.Vacas_ND;
                item.Vacas_A = item.Vacas_A == 0 ? null : item.Vacas_A;
                item.Vacas_Vivos_H = item.Vacas_Vivos_H == 0 ? null : item.Vacas_Vivos_H;
                item.Vacas_Vivos_M = item.Vacas_Vivos_M == 0 ? null : item.Vacas_Vivos_M;

                item.SC = item.SC == 0 ? null : item.SC;

                item.Partos_Vaquillas = item.Partos_Vaquillas == 0 ? null : item.Partos_Vaquillas;
                item.Partos_Vacas = item.Partos_Vacas == 0 ? null : item.Partos_Vacas;

                item.Muertas_Dia = item.Muertas_Dia == 0 ? null : item.Muertas_Dia;
                item.Muertas_Noc = item.Muertas_Noc == 0 ? null : item.Muertas_Noc;
                item.Bp_Gem = item.Bp_Gem == 0 ? null : item.Bp_Gem;

                item.Diferencia_Calostro = item.Diferencia_Calostro == 0 ? null : item.Diferencia_Calostro;
                item.Porcentaje_Calostro = item.Porcentaje_Calostro == 0 ? null : item.Porcentaje_Calostro;
                item.Calidad_Calostro = item.Calidad_Calostro == 0 ? null : item.Calidad_Calostro;
            }
            catch { }
        }

        public void QuitarCeros(List<Hoja4> reporte)
        {
            foreach (Hoja4 item in reporte)
            {
                try
                {
                    item.Ubre_MA = item.Ubre_MA == 0 ? null : item.Ubre_MA;
                    item.Ubre_SL = item.Ubre_SL == 0 ? null : item.Ubre_SL;
                    item.Metabolicos_FL = item.Metabolicos_FL == 0 ? null : item.Metabolicos_FL;
                    item.Metabolicos_CET = item.Metabolicos_CET == 0 ? null : item.Metabolicos_CET;
                    item.Locomotores_BE = item.Locomotores_BE == 0 ? null : item.Locomotores_BE;
                    item.Locomotores_TRA = item.Locomotores_TRA == 0 ? null : item.Locomotores_TRA;
                    item.Locomotores_GA = item.Locomotores_GA == 0 ? null : item.Locomotores_GA;
                    item.Digestivos_AC = item.Digestivos_AC == 0 ? null : item.Digestivos_AC;
                    item.Digestivos_ES = item.Digestivos_ES == 0 ? null : item.Digestivos_ES;
                    item.Digestivos_DI = item.Digestivos_DI == 0 ? null : item.Digestivos_DI;
                    item.Digestivos_TI = item.Digestivos_TI == 0 ? null : item.Digestivos_TI;
                    item.Reproductivos_RE = item.Reproductivos_RE == 0 ? null : item.Reproductivos_RE;
                    item.Reproductivos_ME = item.Reproductivos_ME == 0 ? null : item.Reproductivos_ME;
                    item.Reproductivos_PIO = item.Reproductivos_PIO == 0 ? null : item.Reproductivos_PIO;
                    item.Reproductivos_QUI = item.Reproductivos_QUI == 0 ? null : item.Reproductivos_QUI;
                    item.Reproductivos_CS = item.Reproductivos_CS == 0 ? null : item.Reproductivos_CS;
                    item.Respiratorios_Neu = item.Respiratorios_Neu == 0 ? null : item.Respiratorios_Neu;
                    item.Becerras_Neu = item.Becerras_Neu == 0 ? null : item.Becerras_Neu;
                    item.Becerras_Fie = item.Becerras_Fie == 0 ? null : item.Becerras_Fie;
                    item.Becerras_Di = item.Becerras_Di == 0 ? null : item.Becerras_Di;
                    item.Becerras_Conj = item.Becerras_Conj == 0 ? null : item.Becerras_Conj;
                    item.Vacas_Diag = item.Vacas_Diag == 0 ? null : item.Vacas_Diag;
                    item.Vacas_Pren = item.Vacas_Pren == 0 ? null : item.Vacas_Pren;
                    item.Vacas_Porcentaje_Pren = item.Vacas_Porcentaje_Pren == 0 ? null : item.Vacas_Porcentaje_Pren;
                    item.Vacas_Vacias = item.Vacas_Vacias == 0 ? null : item.Vacas_Vacias;
                    item.Vacas_Porcentaje_Vacias = item.Vacas_Porcentaje_Vacias == 0 ? null : item.Vacas_Porcentaje_Vacias;
                    item.Vaquillas_Diag = item.Vaquillas_Diag == 0 ? null : item.Vaquillas_Diag;
                    item.Vaquillas_Pren = item.Vaquillas_Pren == 0 ? null : item.Vaquillas_Pren;
                    item.Vaquillas_Porcentaje_Pren = item.Vaquillas_Porcentaje_Pren == 0 ? null : item.Vaquillas_Porcentaje_Pren;
                    item.Vaquillas_Vacias = item.Vaquillas_Vacias == 0 ? null : item.Vaquillas_Vacias;
                    item.Vaquillas_Porcentaje_Vacias = item.Vaquillas_Porcentaje_Vacias == 0 ? null : item.Vaquillas_Porcentaje_Vacias;
                    item.Abortos_Vaquillas = item.Abortos_Vaquillas == 0 ? null : item.Abortos_Vaquillas;
                    item.Abortos_Vacas = item.Abortos_Vacas == 0 ? null : item.Abortos_Vacas;
                }
                catch { }
            }
        }

        public void QuitarCeros(Hoja4 item)
        {

            try
            {
                item.Ubre_MA = item.Ubre_MA == 0 ? null : item.Ubre_MA;
                item.Ubre_SL = item.Ubre_SL == 0 ? null : item.Ubre_SL;
                item.Metabolicos_FL = item.Metabolicos_FL == 0 ? null : item.Metabolicos_FL;
                item.Metabolicos_CET = item.Metabolicos_CET == 0 ? null : item.Metabolicos_CET;
                item.Locomotores_BE = item.Locomotores_BE == 0 ? null : item.Locomotores_BE;
                item.Locomotores_TRA = item.Locomotores_TRA == 0 ? null : item.Locomotores_TRA;
                item.Locomotores_GA = item.Locomotores_GA == 0 ? null : item.Locomotores_GA;
                item.Digestivos_AC = item.Digestivos_AC == 0 ? null : item.Digestivos_AC;
                item.Digestivos_ES = item.Digestivos_ES == 0 ? null : item.Digestivos_ES;
                item.Digestivos_DI = item.Digestivos_DI == 0 ? null : item.Digestivos_DI;
                item.Digestivos_TI = item.Digestivos_TI == 0 ? null : item.Digestivos_TI;
                item.Reproductivos_RE = item.Reproductivos_RE == 0 ? null : item.Reproductivos_RE;
                item.Reproductivos_ME = item.Reproductivos_ME == 0 ? null : item.Reproductivos_ME;
                item.Reproductivos_PIO = item.Reproductivos_PIO == 0 ? null : item.Reproductivos_PIO;
                item.Reproductivos_QUI = item.Reproductivos_QUI == 0 ? null : item.Reproductivos_QUI;
                item.Reproductivos_CS = item.Reproductivos_CS == 0 ? null : item.Reproductivos_CS;
                item.Respiratorios_Neu = item.Respiratorios_Neu == 0 ? null : item.Respiratorios_Neu;
                item.Becerras_Neu = item.Becerras_Neu == 0 ? null : item.Becerras_Neu;
                item.Becerras_Fie = item.Becerras_Fie == 0 ? null : item.Becerras_Fie;
                item.Becerras_Di = item.Becerras_Di == 0 ? null : item.Becerras_Di;
                item.Becerras_Conj = item.Becerras_Conj == 0 ? null : item.Becerras_Conj;
                item.Vacas_Diag = item.Vacas_Diag == 0 ? null : item.Vacas_Diag;
                item.Vacas_Pren = item.Vacas_Pren == 0 ? null : item.Vacas_Pren;
                item.Vacas_Porcentaje_Pren = item.Vacas_Porcentaje_Pren == 0 ? null : item.Vacas_Porcentaje_Pren;
                item.Vacas_Vacias = item.Vacas_Vacias == 0 ? null : item.Vacas_Vacias;
                item.Vacas_Porcentaje_Vacias = item.Vacas_Porcentaje_Vacias == 0 ? null : item.Vacas_Porcentaje_Vacias;
                item.Vaquillas_Diag = item.Vaquillas_Diag == 0 ? null : item.Vaquillas_Diag;
                item.Vaquillas_Pren = item.Vaquillas_Pren == 0 ? null : item.Vaquillas_Pren;
                item.Vaquillas_Porcentaje_Pren = item.Vaquillas_Porcentaje_Pren == 0 ? null : item.Vaquillas_Porcentaje_Pren;
                item.Vaquillas_Vacias = item.Vaquillas_Vacias == 0 ? null : item.Vaquillas_Vacias;
                item.Vaquillas_Porcentaje_Vacias = item.Vaquillas_Porcentaje_Vacias == 0 ? null : item.Vaquillas_Porcentaje_Vacias;
                item.Abortos_Vaquillas = item.Abortos_Vaquillas == 0 ? null : item.Abortos_Vaquillas;
                item.Abortos_Vacas = item.Abortos_Vacas == 0 ? null : item.Abortos_Vacas;
            }
            catch { }

        }
        #endregion


        #region Log Errores
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
