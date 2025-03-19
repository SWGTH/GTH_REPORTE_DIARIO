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


namespace ReportePeriodo.Model
{
    public class Model1 : IModel
    {
        string _sqlConnection;
        string _oleDbConnection;
        string _fdbConnection;

        decimal? _precioLeche;
        decimal? _precioLecheAñoAnt;
        List<CalostroYOrdeña> _datosCalostro;

        gth.ConsumoModel consumoModel;

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


        public Model1(string oleDbConnection, string sqlConnection, string fbdConnection)
        {
            this._oleDbConnection = oleDbConnection;
            this._sqlConnection = sqlConnection;
            this._fdbConnection = fbdConnection;
            this.consumoModel = new gth.ConsumoModel(sqlConnection);

            _datosCalostro = new List<CalostroYOrdeña>();
                 
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

        public decimal? PrecioLeche
        {
            get
            {
                if (_precioLeche == null)
                    _precioLeche = 0;

                return _precioLeche;
            }
            set { _precioLeche = value; }
        }

        public decimal? PrecioLecheAñoAnterior
        {
            get
            {
                if (_precioLecheAñoAnt == null)
                    _precioLecheAñoAnt = 0;

                return _precioLecheAñoAnt;
            }
            set { _precioLeche = value; }
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
            catch(DbException ex)
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
            catch(DbException ex)
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
            catch(DbException ex)
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

        public void CargarDatosTeoricos(gthEntity.CentroTyp centro, DateTime fechaInicio, DateTime fechaFin)
        {
            List<gthEntity.CentroTyp> aux = new List<gthEntity.CentroTyp>();
            aux.Add(centro);

            indicadorProduccion = consumoModel.GetIndicadorReporteTeorico(aux, etapaProduccion, fechaInicio, fechaFin);
            indicadorSecas = consumoModel.GetIndicadorReporteTeorico(aux, etapaSecas, fechaInicio, fechaFin);
            indicadorReto = consumoModel.GetIndicadorReporteTeorico(aux, etapaReto, fechaInicio, fechaFin);
            indicadorCrecimiento = consumoModel.GetIndicadorReporteTeorico(aux, etapaCrecimiento, fechaInicio, fechaFin);
            indicadorDesarrollo = consumoModel.GetIndicadorReporteTeorico(aux, etapaDesarrollo, fechaInicio, fechaFin);
            indicadorVaquillas = consumoModel.GetIndicadorReporteTeorico(aux, etapaVaquillas, fechaInicio, fechaFin);

        }

        public void CargarPromediosDatosAlimentacion(gthEntity.CentroTyp centro, DateTime fechaInicio, DateTime fechaFin)
        {
            List<gthEntity.CentroTyp> aux = new List<gthEntity.CentroTyp>();
            aux.Add(centro);

            promedioProduccion = consumoModel.GetIndicadorReportePeriodo(aux, etapaProduccion, fechaInicio, fechaFin);
            promedioSecas = consumoModel.GetIndicadorReportePeriodo(aux, etapaSecas, fechaInicio, fechaFin);
            promedioReto = consumoModel.GetIndicadorReportePeriodo(aux, etapaReto, fechaInicio, fechaFin);
            promedioJaulas = consumoModel.GetIndicadorReportePeriodo(aux, etapaJaulas, fechaInicio, fechaFin);
            promedioCrecimiento = consumoModel.GetIndicadorReportePeriodo(aux, etapaCrecimiento, fechaInicio, fechaFin);
            promedioDesarrollo = consumoModel.GetIndicadorReportePeriodo(aux, etapaDesarrollo, fechaInicio, fechaFin);
            promedioVaquillas = consumoModel.GetIndicadorReportePeriodo(aux, etapaVaquillas, fechaInicio, fechaFin);

            promedioAnteriorProduccion = consumoModel.GetIndicadorReportePeriodo(aux, etapaProduccion, fechaInicio.AddYears(-1), fechaFin.AddYears(-1));
            promedioAnteriorSecas = consumoModel.GetIndicadorReportePeriodo(aux, etapaSecas, fechaInicio.AddYears(-1), fechaFin.AddYears(-1));
            promedioAnteriorReto = consumoModel.GetIndicadorReportePeriodo(aux, etapaReto, fechaInicio.AddYears(-1), fechaFin.AddYears(-1));
            promedioAnteriorJaulas = consumoModel.GetIndicadorReportePeriodo(aux, etapaJaulas, fechaInicio.AddYears(-1), fechaFin.AddYears(-1));
            promedioAnteriorCrecimiento = consumoModel.GetIndicadorReportePeriodo(aux, etapaCrecimiento, fechaInicio.AddYears(-1), fechaFin.AddYears(-1));
            promedioAnteriorDesarrollo = consumoModel.GetIndicadorReportePeriodo(aux, etapaDesarrollo, fechaInicio.AddYears(-1), fechaFin.AddYears(-1));
            promedioAnteriorVaquillas = consumoModel.GetIndicadorReportePeriodo(aux, etapaVaquillas, fechaInicio.AddYears(-1), fechaFin.AddYears(-1));

        }

        public void CargarCalostroOrdeño(DateTime fechaInicio, DateTime fechaFin)
        {
            _datosCalostro = new List<CalostroYOrdeña>();
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
                _datosCalostro = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new CalostroYOrdeña()
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
                InsLog("ReporteDiario", "CargarCalostroOrdeño", $"Error in {nameof(CargarCalostroOrdeño)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }
            
        }
        #endregion

        #region Hoja 1
        public List<Hoja1> GetHoja1(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, List<CalostroYOrdeña> datosCalostro, ref List<decimal?> datosDEC)
        {
            List<Hoja1> response = new List<Hoja1>();            
            datosDEC = new List<decimal?>();

            try
            {
                List<DateTime> fechas = GetFechas(fechaFin);
                List<DatosProduccion> datosProduccion =  GetDatosDeProduccion(fechaInicio, fechaFin);
                List<Mproduc> datosMProduc = rancho.NoIdReal ? GetMediaProduccionNoIdReal(fechaInicio, fechaFin) : GetMediaProduccionNoId(fechaInicio, fechaFin);
                List<NivelCribas> datosCribas = GetNivelCribas(fechaInicio, fechaFin);
                List<Inventario> datosInventario = GetInventario(fechaInicio, fechaFin);

                foreach(DateTime fecha in fechas)
                {
                    Hoja1 hoja = new Hoja1();
                    hoja.Dia = fecha.Day.ToString();

                    DatosProduccion busquedaDatosProduccion = (from x in datosProduccion where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    Mproduc busquedaMProduc = (from x in datosMProduc where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    NivelCribas busquedaCribas = (from x in datosCribas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    Inventario busquedaInventario = (from x in datosInventario where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    CalostroYOrdeña busquedaCalostro = (from x in _datosCalostro where x.Fecha == fecha select x).ToList().FirstOrDefault();

                    hoja.Ordeño = busquedaMProduc.Vacas_Ordeño ?? null;
                    hoja.Secas = busquedaMProduc.Vacas_Secas ?? null;
                    hoja.Hato = busquedaMProduc.Vacas_Hato ?? null;
                    hoja.Porcentaje_Lact = busquedaMProduc.Lactosa ?? null;
                    hoja.Porcentaje_Prot =  busquedaMProduc.Proteina ?? null;
                    hoja.Urea =  busquedaDatosProduccion.Urea ?? null;
                    hoja.Porcentaje_Grasa = busquedaDatosProduccion.Grasa ?? null;
                    hoja.CCS = busquedaDatosProduccion.CCS ?? null;
                    hoja.CTD = busquedaDatosProduccion.CTD ?? null;

                    hoja.Leche = busquedaMProduc.Leche_Producida ?? null;
                    hoja.Antib = busquedaMProduc.Antib ?? null;
                    hoja.Media = busquedaMProduc.Media ?? null;
                    hoja.Total = busquedaMProduc.Total_Producido ?? null;
                    hoja.DEL = busquedaInventario.DelOrd ?? null;
                    hoja.ANT = busquedaMProduc.Vaca_Antib ?? null;

                    hoja.EA = busquedaMProduc.EA ?? null;
                    hoja.ILCA = busquedaMProduc.ILCA_Produccion ?? null;
                    hoja.IC = busquedaMProduc.IC_Produccion ?? null;
                    hoja.Costo_Litro = busquedaMProduc != null ? busquedaMProduc.Media != null && busquedaMProduc.Costo != null ? busquedaMProduc.Media != 0 && busquedaMProduc.Costo != 0 ? busquedaMProduc.Costo / busquedaMProduc.Media : null : null : null;
                    hoja.MH = busquedaMProduc.MH ?? null;
                    hoja.MS = busquedaMProduc.MS ?? null;
                    hoja.Porcentaje_MS = hoja.MH != null && hoja.MS != null ? hoja.MH != 0 ? hoja.MS / hoja.MH : null : null;
                    hoja.SA = busquedaMProduc.SA ?? null;
                    hoja.MSS = busquedaMProduc.MSS ?? null;
                    hoja.EAS = busquedaMProduc.EAS ?? null;
                    hoja.Porcentaje_Sob = hoja.SA != null && hoja.MH != null ? hoja.SA != 0 && hoja.MH != 0 ? hoja.SA / hoja.MH * 100 : null : null;
                    hoja.Costo_Prod = busquedaMProduc.Costo ?? null;
                    hoja.Costo_MS = hoja.Costo_Prod != null && hoja.MS != null ? hoja.Costo_Prod != 0 && hoja.MS != 0 ? hoja.Costo_Prod / hoja.MS : null : null;

                    hoja.SES1 = busquedaCalostro != null ? rancho.VerLitros ? busquedaCalostro.Litros_Sesion1 : busquedaCalostro.Horario_Sesion1 : null;
                    hoja.SES2 = busquedaCalostro != null ? rancho.VerLitros ? busquedaCalostro.Litros_Sesion2 : busquedaCalostro.Horario_Sesion2 : null;
                    hoja.SES3 = busquedaCalostro != null ? rancho.VerLitros ? busquedaCalostro.Litros_Sesion3 : busquedaCalostro.Horario_Sesion3 : null;

                    hoja.Porcentaje_Revueltas = busquedaCalostro.Porcentaje_Revueltas ?? null;

                    response.Add(hoja);
                }

                List<DatosProduccion> busquedaDec1 = (from x in datosProduccion where x.Fecha.Day < 11 select x).ToList();
                List<DatosProduccion> busquedaDec1Dif0 = (from x in datosProduccion where x.Fecha.Day < 11 && x.Grasa != 0 select x).ToList();

                decimal? decPorcGrasa = 0; decimal? valordecPorcGrasa = 0;
                foreach (DatosProduccion itemDec in busquedaDec1)
                    decPorcGrasa += itemDec.Grasa;

                valordecPorcGrasa = busquedaDec1Dif0.Count > 0 ? decPorcGrasa / busquedaDec1Dif0.Count : 0;
                datosDEC.Add(valordecPorcGrasa);

                if (fechaFin.Day > 10)
                {
                    decPorcGrasa = 0; valordecPorcGrasa = 0;
                    List<DatosProduccion> busquedaDec2 = (from x in datosProduccion where x.Fecha.Day > 10 && x.Fecha.Day < 21 select x).ToList();
                    List<DatosProduccion> busquedaDec2Dif0 = (from x in datosProduccion where x.Fecha.Day > 10 && x.Fecha.Day < 21 && x.Grasa != 0 select x).ToList();

                    foreach (DatosProduccion itemDec in busquedaDec2)                    
                        decPorcGrasa += itemDec.Grasa;
                    
                    valordecPorcGrasa = busquedaDec2Dif0.Count > 0 ? decPorcGrasa / busquedaDec2Dif0.Count : 0;
                    datosDEC.Add(valordecPorcGrasa);
                }
                else
                {
                    datosDEC.Add(0);
                }


                if (fechaFin.Day > 20)
                {
                    decPorcGrasa = 0; valordecPorcGrasa = 0;
                    List<DatosProduccion> busquedaDec3 = (from x in datosProduccion where x.Fecha.Day > 21 select x).ToList();
                    List<DatosProduccion> busquedaDec3Dif0 = (from x in datosProduccion where x.Fecha.Day > 21 && x.Grasa != 0 select x).ToList();

                    foreach (DatosProduccion itemDec in busquedaDec3)
                    {
                        decPorcGrasa += itemDec.Grasa;
                    }
                    valordecPorcGrasa = busquedaDec3Dif0.Count > 0 ? decPorcGrasa / busquedaDec3Dif0.Count : 0;
                    datosDEC.Add(valordecPorcGrasa);
                }
                else
                {
                    datosDEC.Add(0);
                }
            }
            catch(Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "GetHoja1", $"Error in {nameof(GetHoja1)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }

            return response;
        }

        public Hoja1 GetPromedioHoja1(Rancho rancho, gthEntity.IndicadorReportePeriodoTyp indicadorOrdeño, DateTime fechaInicio, DateTime fechaFin)
        {
            Hoja1 response = new Hoja1() { Dia = "PROM" };

            try
            {
                CalostroYOrdeña promedioCalostro = GetPromedioCalostro(fechaInicio, fechaFin);
                DatosProduccion promedioDatos = GetTotalDatosDeProduccion(fechaInicio, fechaFin);
                PromedioHoja1 promedioHoja1 = rancho.NoIdReal ? GetPromedioHoja1NoIdReal(fechaInicio, fechaFin) : GetPromedioHoja1NoId(fechaInicio, fechaFin);

                response.Ordeño = promedioHoja1.Ordeño;
                response.Secas = promedioHoja1.Secas;
                response.Hato = promedioHoja1.Hato;
                response.Porcentaje_Lact = promedioHoja1.Porcentaje_Lact;
                response.Porcentaje_Prot = promedioHoja1.Porcentaje_Prot;
                response.Urea = promedioDatos.Urea;
                response.Porcentaje_Grasa = promedioDatos.Grasa;
                response.CCS = promedioDatos.CCS;
                response.CTD = promedioDatos.CTD;
                response.DEC = promedioDatos.Grasa;

                response.Leche = promedioHoja1.Leche;
                response.Antib = promedioHoja1.Antib;
                response.Media = indicadorOrdeño.Media;
                response.Total = promedioHoja1.Total;
                response.DEL = promedioHoja1.DEL;
                response.ANT = promedioHoja1.Ant;

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

                response.Cribas_N1 = promedioHoja1.Cribas_N1;
                response.Cribas_N2 = promedioHoja1.Cribas_N2;
                response.Cribas_N3 = promedioHoja1.Cribas_N3;
                response.Cribas_N4 = promedioHoja1.Cribas_N4;

                response.NoID_Ses1 = promedioHoja1.Sesion1;
                response.NoID_Ses2 = promedioHoja1.Sesion2;
                response.NoID_Ses3 = promedioHoja1.Sesion3;

                response.SES1 = rancho.VerLitros ? promedioCalostro.Litros_Sesion1 : promedioCalostro.Horario_Sesion1;
                response.SES2 = rancho.VerLitros ? promedioCalostro.Litros_Sesion2 : promedioCalostro.Horario_Sesion2;
                response.SES3 = rancho.VerLitros ? promedioCalostro.Litros_Sesion3 : promedioCalostro.Horario_Sesion3;

                response.Porcentaje_Revueltas = promedioCalostro.Porcentaje_Revueltas;
            }
            catch (Exception ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "GetPromedioHoja1", $"Error in {nameof(GetPromedioHoja1)}", ex.Message, ex.Message, ex.StackTrace, guid);
            }

            return response;
        }

        public Hoja1 GetDiferenciaHoja1(Hoja1 promedio, Hoja1 añoAnterior)
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
                InsLog("ReporteDiario", "GetPromedioHoja1", $"Error in {nameof(GetDiferenciaHoja1)}", ex.Message, ex.Message, ex.StackTrace, guid);
            }

            return diferencia;
        }

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
            catch(DbException ex)
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
            catch(DbException ex)
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
            catch (DbException ex)
            {
                Guid guid;
                guid = Guid.NewGuid();
                InsLog("ReporteDiario", "GetPromedioHoja1NoIdReal", $"Error in {nameof(GetPromedioHoja1NoIdReal)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
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
                InsLog("ReporteDiario", "GetPromedioHoja1NoId", $"Error in {nameof(GetPromedioHoja1NoId)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
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
                InsLog("ReporteDiario", "GetPromedioCalostro", $"Error in {nameof(GetPromedioCalostro)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
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
                InsLog("ReporteDiario", "GetTotalDatosDeProduccion", $"Error in {nameof(GetTotalDatosDeProduccion)} with query", ex.Message, ex.Message, ex.StackTrace, guid);
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }


        #endregion

        #region Hoja 2

        #endregion

        #region Hoja 3

        #endregion

        #region Hoja 4

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
            catch(DbException ex)
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
