﻿using System;
using System.Collections.Generic;
using System.Linq;
using ReportePeriodo.Entidad;
using ReportePeriodo.Datos;
using System.Data;
using gth = LibreriaAlimentacion.Entidad;
using LibreriaAlimentacion;
using ght001766q;
using ght001766q.StrongTypesNS;

namespace ReportePeriodo.Modelo
{
    public class Model : IModel
    {
        string conexionAccess;
        string conexionSQL;
        string conexionFDB;
        string conexionSIE;

        List<gth.IndicadorTeorico> indicadorProduccion;
        List<gth.IndicadorTeorico> indicadorSecas;
        List<gth.IndicadorTeorico> indicadorReto;
        List<gth.IndicadorTeorico> indicadorCrecimiento;
        List<gth.IndicadorTeorico> indicadorDesarrollo;
        List<gth.IndicadorTeorico> indicadorVaquillas;

        gth.IndicadorReportePeriodo promedioProduccion;
        gth.IndicadorReportePeriodo promedioSecas;
        gth.IndicadorReportePeriodo promedioReto;
        gth.IndicadorReportePeriodo promedioJaulas;
        gth.IndicadorReportePeriodo promedioCrecimiento;
        gth.IndicadorReportePeriodo promedioDesarrollo;
        gth.IndicadorReportePeriodo promedioVaquillas;

        gth.IndicadorReportePeriodo promedioAntProduccion;
        gth.IndicadorReportePeriodo promedioAntSecas;
        gth.IndicadorReportePeriodo promedioAntReto;
        gth.IndicadorReportePeriodo promedioAntJaulas;
        gth.IndicadorReportePeriodo promedioAntCrecimiento;
        gth.IndicadorReportePeriodo promedioAntDesarrollo;
        gth.IndicadorReportePeriodo promedioAntVaquillas;

        decimal precioLeche;

        public Model(string conexionAccess, string conexionSQL)
        {
            this.conexionAccess = conexionAccess;
            this.conexionSQL = conexionSQL;

            indicadorProduccion = new List<gth.IndicadorTeorico>();
            indicadorSecas = new List<gth.IndicadorTeorico>();
            indicadorReto = new List<gth.IndicadorTeorico>();
            indicadorCrecimiento = new List<gth.IndicadorTeorico>();
            indicadorDesarrollo = new List<gth.IndicadorTeorico>();
            indicadorVaquillas = new List<gth.IndicadorTeorico>();

            promedioProduccion = new gth.IndicadorReportePeriodo();
            promedioSecas = new gth.IndicadorReportePeriodo();
            promedioReto = new gth.IndicadorReportePeriodo();
            promedioJaulas = new gth.IndicadorReportePeriodo();
            promedioCrecimiento = new gth.IndicadorReportePeriodo();
            promedioDesarrollo = new gth.IndicadorReportePeriodo();
            promedioVaquillas = new gth.IndicadorReportePeriodo();

            promedioAntProduccion = new gth.IndicadorReportePeriodo();
            promedioAntSecas = new gth.IndicadorReportePeriodo();
            promedioAntReto = new gth.IndicadorReportePeriodo();
            promedioAntJaulas = new gth.IndicadorReportePeriodo();
            promedioAntCrecimiento = new gth.IndicadorReportePeriodo();
            promedioAntDesarrollo = new gth.IndicadorReportePeriodo();
            promedioAntVaquillas = new gth.IndicadorReportePeriodo();

            precioLeche = 0;
        }

        #region Propiedades

        public string ConexionAccess
        {
            get { return conexionAccess; }
            set { conexionAccess = value; }
        }

        public string ConexionSQL
        {
            get { return conexionSQL; }
            set { conexionSQL = value; }
        }

        public string ConexionFDB
        {
            get { return conexionFDB; }
            set { conexionFDB = value; }
        }

        public string ConexionSIE
        {
            get { return conexionSIE; }
            set { conexionSIE = value; }
        }

        public List<gth.IndicadorTeorico> IndicadorProduccion
        {
            get { return indicadorProduccion; }
            set { indicadorProduccion = value; }
        }

        public List<gth.IndicadorTeorico> IndicadorSecas
        {
            get { return indicadorSecas; }
            set { indicadorSecas = value; }
        }

        public List<gth.IndicadorTeorico> IndicadorReto
        {
            get { return indicadorReto; }
            set { indicadorReto = value; }
        }

        public List<gth.IndicadorTeorico> IndicadorCrecimiento
        {
            get { return indicadorCrecimiento; }
            set { indicadorCrecimiento = value; }
        }

        public List<gth.IndicadorTeorico> IndicadorDesarrollo
        {
            get { return indicadorDesarrollo; }
            set { indicadorDesarrollo = value; }
        }

        public List<gth.IndicadorTeorico> IndicadorVaquillas
        {
            get { return indicadorVaquillas; }
            set { indicadorVaquillas = value; }
        }

        public gth.IndicadorReportePeriodo PromedioProduccion
        {
            get { return promedioProduccion; }
            set { promedioProduccion = value; }
        }

        public gth.IndicadorReportePeriodo PromedioSecas
        {
            get { return promedioSecas; }
            set { promedioSecas = value; }
        }

        public gth.IndicadorReportePeriodo PromedioReto
        {
            get { return promedioReto; }
            set { promedioReto = value; }
        }

        public gth.IndicadorReportePeriodo PromedioJaulas
        {
            get { return promedioJaulas; }
            set { promedioJaulas = value; }
        }

        public gth.IndicadorReportePeriodo PromedioCrecimiento
        {
            get { return promedioCrecimiento; }
            set { promedioCrecimiento = value; }
        }

        public gth.IndicadorReportePeriodo PromedioDesarrollo
        {
            get { return promedioDesarrollo; }
            set { promedioDesarrollo = value; }
        }

        public gth.IndicadorReportePeriodo PromedioVaquillas
        {
            get { return promedioVaquillas; }
            set { promedioVaquillas = value; }
        }

        public gth.IndicadorReportePeriodo PromedioAntProduccion
        {
            get { return promedioAntProduccion; }
            set { promedioAntProduccion = value; }
        }

        public gth.IndicadorReportePeriodo PromedioAntSecas
        {
            get { return promedioAntSecas; }
            set { promedioAntSecas = value; }
        }

        public gth.IndicadorReportePeriodo PromedioAntReto
        {
            get { return promedioAntReto; }
            set { promedioAntReto = value; }
        }

        public gth.IndicadorReportePeriodo PromedioAntJaulas
        {
            get { return promedioAntJaulas; }
            set { promedioAntJaulas = value; }

        }

        public gth.IndicadorReportePeriodo PromedioAntCrecimiento
        {
            get { return promedioAntCrecimiento; }
            set { promedioAntCrecimiento = value; }
        }

        public gth.IndicadorReportePeriodo PromedioAntDesarrollo
        {
            get { return promedioAntDesarrollo; }
            set { promedioAntDesarrollo = value; }
        }

        public gth.IndicadorReportePeriodo PromedioAntVaquillas
        {
            get { return promedioAntVaquillas; }
            set { promedioAntVaquillas = value; }
        }

        public decimal PrecioLeche
        {
            get { return precioLeche; }
            set { precioLeche = value; }
        }

        #endregion

        #region Metodos publicos
        public Rancho Rancho(ref string mensaje)
        {
            Rancho rancho = new Rancho();
            mensaje = string.Empty;

            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);

            try
            {
                string query = @"SELECT rl.RanchoLocal, r.Nombre, r.Empresa, r.NOIDREAL, r.VERLITROS, r.pesadores
                                 FROM rancholocal rl
                                 LEFT JOIN ranchos r on rl.rancholocal = r.Clave";

                db.Conectar();
                db.CrearComando(query, tipoComandoAccess.query);
                rancho = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new Rancho()
                {
                    Ran_ID = x["RanchoLocal"] != DBNull.Value ? Convert.ToInt32(x["RanchoLocal"]) : 0,
                    Ran_Nombre = x["Nombre"] != DBNull.Value ? x["Nombre"].ToString() : string.Empty,
                    Emp_ID = x["Empresa"] != DBNull.Value ? Convert.ToInt32(x["Empresa"]) : 0,
                    No_ID_Real = x["NOIDREAL"] != DBNull.Value ? Convert.ToBoolean(x["NOIDREAL"]) : false,
                    Ver_Litros = x["VERLITROS"] != DBNull.Value ? Convert.ToBoolean(x["VERLITROS"]) : false,
                    Pesadores = x["pesadores"] != DBNull.Value ? Convert.ToBoolean(x["pesadores"]) : false

                }).ToList().FirstOrDefault();
            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return rancho;
        }

        public int HoraCorte(ref string mensaje)
        {
            int hora = 0;
            ModeloDatosFDB db = new ModeloDatosFDB(conexionFDB);
            mensaje = string.Empty;

            try
            {
                string query = @"SELECT PARAMVALUE AS Hora
                                 FROM BEDRIJF_PARAMS bp 
                                 WHERE NAME = 'DSTimeShift'";

                db.Conectar();
                db.CrearComando(query, tipoComandoFDB.query);
                DataTable dt = db.EjecutarConsultaTabla();
                if (dt.Rows.Count > 0)
                    hora = dt.Rows[0][0] != DBNull.Value ? Convert.ToInt32(dt.Rows[0][0]) : 0;
            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return hora;
        }

        public string Erp_Clave(int ranID, ref string mensaje)
        {
            string erp = "0227";
            ModeloDatosSQL db = new ModeloDatosSQL(conexionSQL);
            mensaje = string.Empty;

            try
            {
                string query = @"SELECT erp_id
                                 FROM DBSIO.dbo.configuracion
                                 where ran_id = @rancho";

                db.Conectar();
                db.CrearComando(query, tipoComandoSQL.query);
                db.AsignarParametro("@rancho", ranID);
                DataTable dt = db.EjecutarConsultaTabla();

                if (dt.Rows.Count > 0)
                    erp = dt.Rows[0][0] != DBNull.Value ? dt.Rows[0][0].ToString() : "0227";

            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }


            return erp;
        }

        public List<Hoja1> ReporteHoja1(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, out List<CalostroYOrdeña> datosCalostro, out List<decimal?> listaDEC, ref string mensaje)
        {
            List<Hoja1> response = new List<Hoja1>();
            mensaje = string.Empty;
            datosCalostro = new List<CalostroYOrdeña>();
            listaDEC = new List<decimal?>();

            try
            {
                List<DateTime> fechasReporte = ListaFechasReporte(fechaFin);
                List<DatosProduccion> datosProduccion = DatosDeProduccion(fechaInicio, fechaFin);
                List<Mproduc> datosMProduc = rancho.No_ID_Real ? MediaProduccionNoIdReal(fechaInicio, fechaFin) : MediaProduccionNoId(fechaInicio, fechaFin);
                List<Cribas> datosCribas = NivelCriba(fechaInicio, fechaFin);
                datosCalostro = CalostroOrdeño(fechaInicio, fechaFin);
                List<Inventario> datosInventario = Inventario(fechaInicio, fechaFin);

                foreach (DateTime fecha in fechasReporte)
                {
                    Hoja1 hoja = new Hoja1();
                    hoja.Dia = fecha.Day.ToString();

                    DatosProduccion busquedaDatosProduccion = (from x in datosProduccion where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    Mproduc busquedaMProduc = (from x in datosMProduc where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    Cribas busquedaCribas = (from x in datosCribas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    CalostroYOrdeña busquedaCalostro = (from x in datosCalostro where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    Inventario busquedaInventario = (from x in datosInventario where x.Fecha == fecha select x).ToList().FirstOrDefault();

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

                    hoja.SES1 = busquedaCalostro != null ? rancho.Ver_Litros ? busquedaCalostro.Litros_Sesion1 : busquedaCalostro.Horario_Sesion1 : null;
                    hoja.SES2 = busquedaCalostro != null ? rancho.Ver_Litros ? busquedaCalostro.Litros_Sesion2 : busquedaCalostro.Horario_Sesion2 : null;
                    hoja.SES3 = busquedaCalostro != null ? rancho.Ver_Litros ? busquedaCalostro.Litros_Sesion3 : busquedaCalostro.Horario_Sesion3 : null;

                    hoja.Porcentaje_Revueltas = busquedaCalostro != null ? busquedaCalostro.Porcentaje_Revueltas : null;

                    response.Add(hoja);
                }

                List<DatosProduccion> busquedaDec1 = (from x in datosProduccion where x.Fecha.Day < 11 select x).ToList();
                List<DatosProduccion> busquedaDec1Dif0 = (from x in datosProduccion where x.Fecha.Day < 11 && x.Grasa != 0 select x).ToList();

                decimal? decPorcGrasa = 0; decimal? valordecPorcGrasa = 0;
                foreach (DatosProduccion itemDec in busquedaDec1)
                    decPorcGrasa += itemDec.Grasa;

                valordecPorcGrasa = busquedaDec1Dif0.Count > 0 ? decPorcGrasa / busquedaDec1Dif0.Count : 0;
                listaDEC.Add(valordecPorcGrasa);

                if (fechaFin.Day > 10)
                {
                    decPorcGrasa = 0; valordecPorcGrasa = 0;
                    List<DatosProduccion> busquedaDec2 = (from x in datosProduccion where x.Fecha.Day > 10 && x.Fecha.Day < 21 select x).ToList();
                    List<DatosProduccion> busquedaDec2Dif0 = (from x in datosProduccion where x.Fecha.Day > 10 && x.Fecha.Day < 21 && x.Grasa != 0 select x).ToList();

                    foreach (DatosProduccion itemDec in busquedaDec2)
                    {
                        decPorcGrasa += itemDec.Grasa;
                    }
                    valordecPorcGrasa = busquedaDec2Dif0.Count > 0 ? decPorcGrasa / busquedaDec2Dif0.Count : 0;
                    listaDEC.Add(valordecPorcGrasa);
                }
                else
                {
                    listaDEC.Add(0);
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
                    listaDEC.Add(valordecPorcGrasa);
                }
                else
                {
                    listaDEC.Add(0);
                }

            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }

            return response;
        }

        public void CargarDatosTeoricos(Rancho rancho, DateTime fechaFin, ref string mensaje)
        {
            GTH.IndicadoresTeoricosxEstablo(rancho.Ran_ID, "10,11,12,13,91", fechaFin, out indicadorProduccion, ref mensaje);
            GTH.IndicadoresTeoricosxEstablo(rancho.Ran_ID, "21,92", fechaFin, out indicadorSecas, ref mensaje);
            GTH.IndicadoresTeoricosxEstablo(rancho.Ran_ID, "22,94", fechaFin, out indicadorReto, ref mensaje);
            GTH.IndicadoresTeoricosxEstablo(rancho.Ran_ID, "32,96", fechaFin, out indicadorCrecimiento, ref mensaje);
            GTH.IndicadoresTeoricosxEstablo(rancho.Ran_ID, "33,97", fechaFin, out indicadorDesarrollo, ref mensaje);
            GTH.IndicadoresTeoricosxEstablo(rancho.Ran_ID, "34,98", fechaFin, out indicadorVaquillas, ref mensaje);
        }

        public void CargarPromediosDatosAlimentacion(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            mensaje = string.Empty;

            try
            {
                List<gth.ReportePeriodo> ingredientesAux = new List<gth.ReportePeriodo>();
                gth.ReportePeriodo sobranteAux = new gth.ReportePeriodo();

                GTH.ReportePeriodo(rancho.Ran_ID.ToString(), rancho.TimeShiftTracker, "10,11,12,13", fechaInicio, fechaFin, out ingredientesAux, out promedioProduccion, out sobranteAux);
                GTH.ReportePeriodo(rancho.Ran_ID.ToString(), rancho.TimeShiftTracker, "21", fechaInicio, fechaFin, out ingredientesAux, out promedioSecas, out sobranteAux);
                GTH.ReportePeriodo(rancho.Ran_ID.ToString(), rancho.TimeShiftTracker, "22", fechaInicio, fechaFin, out ingredientesAux, out promedioReto, out sobranteAux);
                GTH.ReportePeriodo(rancho.Ran_ID.ToString(), rancho.TimeShiftTracker, "31", fechaInicio, fechaFin, out ingredientesAux, out promedioJaulas, out sobranteAux);
                GTH.ReportePeriodo(rancho.Ran_ID.ToString(), rancho.TimeShiftTracker, "32", fechaInicio, fechaFin, out ingredientesAux, out promedioCrecimiento, out sobranteAux);
                GTH.ReportePeriodo(rancho.Ran_ID.ToString(), rancho.TimeShiftTracker, "33", fechaInicio, fechaFin, out ingredientesAux, out promedioDesarrollo, out sobranteAux);
                GTH.ReportePeriodo(rancho.Ran_ID.ToString(), rancho.TimeShiftTracker, "34", fechaInicio, fechaFin, out ingredientesAux, out promedioVaquillas, out sobranteAux);

                GTH.ReportePeriodo(rancho.Ran_ID.ToString(), rancho.TimeShiftTracker, "10,11,12,13", fechaInicio.AddYears(-1), fechaFin.AddYears(-1), out ingredientesAux, out promedioAntProduccion, out sobranteAux);
                GTH.ReportePeriodo(rancho.Ran_ID.ToString(), rancho.TimeShiftTracker, "21", fechaInicio.AddYears(-1), fechaFin.AddYears(-1), out ingredientesAux, out promedioAntSecas, out sobranteAux);
                GTH.ReportePeriodo(rancho.Ran_ID.ToString(), rancho.TimeShiftTracker, "22", fechaInicio.AddYears(-1), fechaFin.AddYears(-1), out ingredientesAux, out promedioAntReto, out sobranteAux);
                GTH.ReportePeriodo(rancho.Ran_ID.ToString(), rancho.TimeShiftTracker, "31", fechaInicio.AddYears(-1), fechaFin.AddYears(-1), out ingredientesAux, out promedioAntJaulas, out sobranteAux);
                GTH.ReportePeriodo(rancho.Ran_ID.ToString(), rancho.TimeShiftTracker, "32", fechaInicio.AddYears(-1), fechaFin.AddYears(-1), out ingredientesAux, out promedioAntCrecimiento, out sobranteAux);
                GTH.ReportePeriodo(rancho.Ran_ID.ToString(), rancho.TimeShiftTracker, "33", fechaInicio.AddYears(-1), fechaFin.AddYears(-1), out ingredientesAux, out promedioAntDesarrollo, out sobranteAux);
                GTH.ReportePeriodo(rancho.Ran_ID.ToString(), rancho.TimeShiftTracker, "34", fechaInicio.AddYears(-1), fechaFin.AddYears(-1), out ingredientesAux, out promedioAntVaquillas, out sobranteAux);

            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
        }

        public Hoja1 PromedioHoja1(Rancho rancho, gth.IndicadorReportePeriodo indicadorOrdeño, DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            Hoja1 response = new Hoja1() { Dia = "PROM" };
            PromedioHoja1 promH1 = rancho.No_ID_Real ? PromedioHoja1NoIdReal(fechaInicio, fechaFin, ref mensaje) : PromedioHoja1NoId(fechaInicio, fechaFin, ref mensaje);
            mensaje = string.Empty;

            try
            {

                CalostroYOrdeña promedioCalostro = PromedioCalostro(fechaInicio, fechaFin);
                DatosProduccion promedioDatos = TotalDatosDeProduccion(fechaInicio, fechaFin);

                response.Ordeño = promH1.Ordeño;
                response.Secas = promH1.Secas;
                response.Hato = promH1.Hato;
                response.Porcentaje_Lact = promH1.Porcentaje_Lact;
                response.Porcentaje_Prot = promH1.Porcentaje_Prot;
                response.Urea = promedioDatos.Urea;
                response.Porcentaje_Grasa = promedioDatos.Grasa;
                response.CCS = promedioDatos.CCS;
                response.CTD = promedioDatos.CTD;
                response.DEC = promedioDatos.Grasa;

                response.Leche = promH1.Leche;
                response.Antib = promH1.Antib;
                response.Media = (decimal)indicadorOrdeño.MEDIA;
                response.Total = promH1.Total;
                response.DEL = promH1.DEL;
                response.ANT = promH1.Ant;

                response.EA = (decimal)indicadorOrdeño.EA;
                response.ILCA = (decimal)indicadorOrdeño.ILCA_PRODUCCION;
                response.IC = (decimal)indicadorOrdeño.IC_PRODUCCION;
                response.Costo_Litro = (decimal)indicadorOrdeño.PRECIOL;
                response.MH = (decimal)indicadorOrdeño.MH;
                response.Porcentaje_MS = (decimal)(indicadorOrdeño.PORCENTAJEMS / 100);
                response.MS = (decimal)indicadorOrdeño.MS;
                response.SA = (decimal)indicadorOrdeño.SA;
                response.MSS = (decimal)indicadorOrdeño.MSS;
                response.EAS = (decimal)indicadorOrdeño.EAS;
                response.Porcentaje_Sob = indicadorOrdeño.MH > 0 ? Convert.ToDecimal(indicadorOrdeño.SA / indicadorOrdeño.MH * 100) : 0;
                response.Costo_Prod = (decimal)indicadorOrdeño.COSTO;
                response.Costo_MS = (decimal)indicadorOrdeño.PRECIOKGMS;

                response.Cribas_N1 = promH1.Cribas_N1;
                response.Cribas_N2 = promH1.Cribas_N2;
                response.Cribas_N3 = promH1.Cribas_N3;
                response.Cribas_N4 = promH1.Cribas_N4;

                response.NoID_Ses1 = promH1.Sesion1;
                response.NoID_Ses2 = promH1.Sesion2;
                response.NoID_Ses3 = promH1.Sesion3;

                response.SES1 = rancho.Ver_Litros ? promedioCalostro.Litros_Sesion1 : promedioCalostro.Horario_Sesion1;
                response.SES2 = rancho.Ver_Litros ? promedioCalostro.Litros_Sesion2 : promedioCalostro.Horario_Sesion2;
                response.SES3 = rancho.Ver_Litros ? promedioCalostro.Litros_Sesion3 : promedioCalostro.Horario_Sesion3;

                response.Porcentaje_Revueltas = promedioCalostro.Porcentaje_Revueltas;
            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }

            return response;
        }

        public void AsignarColorimetriaHoja1(Rancho rancho, List<Hoja1> reporte, List<CalostroYOrdeña> datosCalostro, Hoja1 promedio, DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            mensaje = string.Empty;
            try
            {
                CalostroYOrdeña busquedaCalostro = new CalostroYOrdeña();
                List<BacteriologiaLeche> datosBacteriologia = BacteriologiaLeche(fechaInicio, fechaFin, ref mensaje);
                BacteriologiaLeche busquedaBacteriologia = new BacteriologiaLeche();
                gth.IndicadorTeorico busquedaIndicador = new gth.IndicadorTeorico();

                foreach (Hoja1 item in reporte)
                {
                    busquedaCalostro = (from x in datosCalostro where x.Fecha.Day.ToString() == item.Dia select x).ToList().FirstOrDefault();
                    busquedaBacteriologia = (from x in datosBacteriologia where x.FECHA.Day.ToString() == item.Dia select x).ToList().FirstOrDefault();
                    busquedaIndicador = (from x in indicadorProduccion where x.FECHA.Day.ToString() == item.Dia select x).ToList().FirstOrDefault();

                    item.Color_ILCA = Color1Hoja1(item.ILCA, Convert.ToDecimal(promedioProduccion.ILCA_PRODUCCION));
                    item.Color_IC = Color1Hoja1(item.IC, Convert.ToDecimal(promedioProduccion.IC_PRODUCCION));
                    item.Color_CostoLitro = Color2Hoja1(item.Costo_Litro, promedio.Costo_Litro);

                    item.Color_MH = ColorDato(item.Ordeño, item.MH, busquedaIndicador.MH);
                    item.Color_PorcentajeMs = ColorDato(item.Ordeño, item.Porcentaje_MS, busquedaIndicador.PORCENTAJE_MS / 100);
                    item.Color_MS = ColorDato(item.Ordeño, item.MS, busquedaIndicador.MS);
                    item.Color_CostoProd = ColorDato(item.Ordeño, item.Costo_Prod, busquedaIndicador.COSTO);
                    item.Color_CostoMs = ColorDato(item.Ordeño, item.Costo_MS, busquedaIndicador.KGMS);

                    item.Color_Leche = Color1Hoja1(item.Leche, promedio.Leche);
                    item.Color_Media = Color1Hoja1(item.Media, promedio.Media);
                    item.Color_Total = Color1Hoja1(item.Total, promedio.Total);
                    item.Color_Ses1 = ColorSes(rancho.Pesadores, item.Ordeño, item.NoID_Ses1);
                    item.Color_Ses2 = ColorSes(rancho.Pesadores, item.Ordeño, item.NoID_Ses2);
                    item.Color_Ses3 = ColorSes(rancho.Pesadores, item.Ordeño, item.NoID_Ses3);
                    item.Color_CTD = item.CTD != null ? busquedaBacteriologia != null ? busquedaBacteriologia.BACTEROLOGIA < 0.2M ? "#FFC9C9" : "#F2F2F2" : string.Empty : string.Empty;
                    item.Color_ANT = ColorAntHoja1(item.Media, item.Antib, item.ANT);
                    item.Color_HoraSes1 = busquedaCalostro != null ? ColorSesionesHora(busquedaCalostro.Horario_Sesion1) : string.Empty;
                    item.Color_HoraSes2 = busquedaCalostro != null ? ColorSesionesHora(busquedaCalostro.Horario_Sesion2) : string.Empty;
                    item.Color_HoraSes3 = busquedaCalostro != null ? ColorSesionesHora(busquedaCalostro.Horario_Sesion3) : string.Empty;
                    item.Color_Ordeño = busquedaCalostro != null ? ColorSalaOrdeño(busquedaCalostro.Capacidad_Ordeño) : string.Empty;

                    if (item.Color_HoraSes1 != "#FFC9C9" && item.Color_HoraSes1 != string.Empty && item.Color_HoraSes1 != "")
                        item.Color_HoraSes1 = busquedaCalostro != null ? busquedaCalostro.Horario_Dif_Sesion1 < 10 ? "#FFC9C9" : item.Color_HoraSes1 : item.Color_HoraSes1;

                    if (item.Color_HoraSes2 != "#FFC9C9" && item.Color_HoraSes2 != string.Empty && item.Color_HoraSes2 != "")
                        item.Color_HoraSes2 = busquedaCalostro != null ? busquedaCalostro.Horario_Dif_Sesion2 < 10 ? "#FFC9C9" : item.Color_HoraSes2 : item.Color_HoraSes2;

                    if (item.Color_HoraSes3 != "#FFC9C9" && item.Color_HoraSes3 != string.Empty && item.Color_HoraSes3 != "")
                        item.Color_HoraSes3 = busquedaCalostro != null ? busquedaCalostro.Horario_Dif_Sesion3 < 10 ? "#FFC9C9" : item.Color_HoraSes3 : item.Color_HoraSes3;
                }
            }
            catch { }
        }

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
            catch { }

            return diferencia;
        }

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
            catch { }

            return porcentaje;
        }

        public List<Hoja1> EspaciosEnBlancoHoja1(int renglones)
        {
            List<Hoja1> response = new List<Hoja1>();
            int renglonesTotal = 32 - renglones;

            for (int i = 0; i < renglonesTotal; i++)
            {
                response.Add(new Hoja1());
            }

            return response;
        }

        public void QuitarCeros(List<Hoja1> reporte)
        {
            foreach (Hoja1 item in reporte)
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

        public void QuitarCeros(List<Hoja2> reporte)
        {
            foreach (Hoja2 item in reporte)
            {

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

        public void QuitarCeros(Hoja1 item)
        {
            try
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
            }
            catch { }
        }

        public List<Hoja2> ReporteHoja2(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            List<Hoja2> response = new List<Hoja2>();
            mensaje = string.Empty;

            try
            {
                List<DateTime> fechasReporte = ListaFechasReporte(fechaFin);
                List<InventarioAfiXDia> datosInventarioAFI = InventariosAFI(fechaInicio, fechaFin, ref mensaje);
                List<Mproduc2> datosMProduc = MediaProduccion2(fechaInicio, fechaFin, ref mensaje);



                foreach (DateTime fecha in fechasReporte)
                {
                    InventarioAfiXDia busqudaInventarioAFI = (from x in datosInventarioAFI where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    Mproduc2 busquedaMproduc = (from x in datosMProduc where x.Fecha == fecha select x).ToList().FirstOrDefault();

                    Hoja2 item = new Hoja2();
                    item.Dia = fecha.Day.ToString();

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
                    item.IngresoxAnimal = busquedaMproduc != null ? item.Inventario_Total > 0 ? (busquedaMproduc.LecheProd * precioLeche) / item.Inventario_Total : 0 : 0;

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
                mensaje = ex.Message;
            }


            return response;
        }

        public Hoja2 PromedioHoja2(Rancho rancho, gth.IndicadorReportePeriodo indicadorOrdeño, gth.IndicadorReportePeriodo indicadorSecas, gth.IndicadorReportePeriodo indicadorReto,
            gth.IndicadorReportePeriodo indicadorJaulas, gth.IndicadorReportePeriodo indicadorCrecimiento, gth.IndicadorReportePeriodo indicadorDesarrollo,
            gth.IndicadorReportePeriodo indicadorVaquillas, DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            Hoja2 promedio = new Hoja2() { Dia = "PROM" };
            InventarioAfiXDia inventario = PromedioInventariosAFI(fechaInicio, fechaFin, ref mensaje);

            bool esPrecioLecheFacturado = EsPrecioLecheFacturado(fechaFin);
            decimal precioLecheX = esPrecioLecheFacturado ? PrecioLecheFacturado(fechaInicio, fechaFin, ref mensaje) : PrecioLecheCalculado(rancho.Ran_ID, fechaInicio, fechaFin, ref mensaje);
            decimal lecheProducida = LecheProducida(fechaInicio, fechaFin, ref mensaje);

            try
            {
                promedio.Jaulas_Inventario = (decimal)indicadorJaulas.ANIMELES;
                promedio.Jaulas_Costo = (decimal)indicadorJaulas.COSTO;

                promedio.Crecimiento_Inventario = (decimal)indicadorCrecimiento.ANIMELES;
                promedio.Crecimiento_MH = (decimal)indicadorCrecimiento.MH;
                promedio.Crecimiento_MS = (decimal)indicadorCrecimiento.MS;
                promedio.Crecimiento_PorcentajeMS = (decimal)indicadorCrecimiento.PORCENTAJEMS;
                promedio.Crecimiento_Costo = (decimal)indicadorCrecimiento.COSTO;
                promedio.Crecimiento_CostoMS = (decimal)indicadorCrecimiento.PRECIOKGMS;

                promedio.Desarrollo_Inventario = (decimal)indicadorDesarrollo.ANIMELES;
                promedio.Desarrollo_MH = (decimal)indicadorDesarrollo.MH;
                promedio.Desarrollo_MS = (decimal)indicadorDesarrollo.MS;
                promedio.Desarrollo_PorcentajeMS = (decimal)indicadorDesarrollo.PORCENTAJEMS;
                promedio.Desarrollo_Costo = (decimal)indicadorDesarrollo.COSTO;
                promedio.Desarrollo_CostoMS = (decimal)indicadorDesarrollo.PRECIOKGMS;

                promedio.Vaquillas_Inventario = (decimal)indicadorVaquillas.ANIMELES;
                promedio.Vaquillas_MH = (decimal)indicadorVaquillas.MH;
                promedio.Vaquillas_MS = (decimal)indicadorVaquillas.MS;
                promedio.Vaquillas_PorcentajeMS = (decimal)indicadorVaquillas.PORCENTAJEMS;
                promedio.Vaquillas_Costo = (decimal)indicadorVaquillas.COSTO;
                promedio.Vaquillas_CostoMS = (decimal)indicadorVaquillas.PRECIOKGMS;

                promedio.Secas_Inventario = (decimal)indicadorSecas.ANIMELES;
                promedio.Secas_MH = (decimal)indicadorSecas.MH;
                promedio.Secas_PorcentajeMS = (decimal)indicadorSecas.PORCENTAJEMS;
                promedio.Secas_MS = (decimal)indicadorSecas.MS;
                promedio.Secas_SA = (decimal)indicadorSecas.SA;
                promedio.Secas_Mss = (decimal)indicadorSecas.MSS;
                promedio.Secas_PorcentajeSob = indicadorSecas.MH > 0 ? Convert.ToDecimal(indicadorSecas.SA / indicadorSecas.MH * 100) : 0;
                promedio.Secas_Costo = (decimal)indicadorSecas.COSTO;
                promedio.Secas_CostoMS = (decimal)indicadorSecas.PRECIOKGMS;

                promedio.Reto_Inventario = (decimal)indicadorReto.ANIMELES;
                promedio.Reto_MH = (decimal)indicadorReto.MH;
                promedio.Reto_PorcentajeMS = (decimal)indicadorReto.PORCENTAJEMS;
                promedio.Reto_MS = (decimal)indicadorReto.MS;
                promedio.Reto_SA = (decimal)indicadorReto.SA;
                promedio.Reto_Mss = (decimal)indicadorReto.MSS;
                promedio.Reto_PorcentajeSob = indicadorReto.MH > 0 ? Convert.ToDecimal(indicadorReto.SA / indicadorReto.MH * 100) : 0;
                promedio.Reto_Costo = (decimal)indicadorReto.COSTO;
                promedio.Reto_CostoMS = (decimal)indicadorReto.PRECIOKGMS;

                promedio.Inventario_Total = inventario.InventarioTotal;

                decimal ixA = 0;
                try { ixA = promedio.Inventario_Total > 0 ? Convert.ToDecimal((lecheProducida * precioLecheX) / promedio.Inventario_Total) : 0; }
                catch { }

                promedio.IngresoxAnimal = Math.Round(ixA, 5);

                decimal? auxJaulas = promedio.Jaulas_Costo * promedio.Jaulas_Inventario;
                decimal? auxCrecimiento = promedio.Crecimiento_Costo * promedio.Crecimiento_Inventario;
                decimal? auxDesarrollo = promedio.Desarrollo_Costo * promedio.Desarrollo_Inventario;
                decimal? auxVaquillas = promedio.Vaquillas_Costo * promedio.Vaquillas_Inventario;
                decimal? auxSecas = promedio.Secas_Costo * promedio.Secas_Inventario;
                decimal? auxReto = promedio.Reto_Costo * promedio.Reto_Inventario;
                decimal? auxOrdeño = Convert.ToDecimal(indicadorOrdeño.ANIMELES * indicadorOrdeño.COSTO);
                decimal? auxCosto = auxJaulas + auxCrecimiento + auxDesarrollo + auxVaquillas + auxSecas + auxReto + auxOrdeño;

                decimal cxa = 0;
                try { cxa = promedio.Inventario_Total > 0 ? Convert.ToDecimal(auxCosto / promedio.Inventario_Total) : 0; }
                catch { }

                promedio.CostoxAnimal = Math.Round(cxa, 5);
                promedio.UtilidadxAnimal = promedio.IngresoxAnimal - promedio.CostoxAnimal;
                promedio.Porcentaje_CostoxAnimal = promedio.IngresoxAnimal > 0 ? promedio.CostoxAnimal / promedio.IngresoxAnimal : 0;
                promedio.Porcentaje_UtilidadxAnimal = promedio.IngresoxAnimal > 0 ? promedio.UtilidadxAnimal / promedio.IngresoxAnimal : 0;

            }
            catch { }

            return promedio;
        }

        public Hoja2 DiferenciaHoja2(Hoja2 promedio, Hoja2 promedioAñoAnt)
        {
            Hoja2 difencia = new Hoja2() { Dia = "DIF #" };

            try
            {
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
            catch { }

            return difencia;
        }

        public Hoja2 PorcentajeDiferenciaHoja2(Hoja2 diferencia, Hoja2 promedioAñoAnt)
        {
            Hoja2 porcentaje = new Hoja2() { Dia = "DIF %" };

            try
            {
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
            catch { }

            return porcentaje;

        }

        public List<Hoja2> EspaciosEnBlancoHoja2(int renglones)
        {
            List<Hoja2> response = new List<Hoja2>();
            int renglonesTotal = 32 - renglones;

            for (int i = 0; i < renglonesTotal; i++)
            {
                response.Add(new Hoja2());
            }

            return response;
        }

        public void AsignarColorimetriaHoja2(List<Hoja2> reporte, Utilidad utilidad)
        {
            gth.IndicadorTeorico busquedaindicadorCrecimiento = new gth.IndicadorTeorico();
            gth.IndicadorTeorico busquedaindicadorDesarrollo = new gth.IndicadorTeorico();
            gth.IndicadorTeorico busquedaindicadorVaquillas = new gth.IndicadorTeorico();
            gth.IndicadorTeorico busquedaindicadorSecas = new gth.IndicadorTeorico();
            gth.IndicadorTeorico busquedaindicadorReto = new gth.IndicadorTeorico();

            try
            {
                foreach (Hoja2 item in reporte)
                {
                    busquedaindicadorCrecimiento = (from x in indicadorCrecimiento where x.FECHA.Day.ToString() == item.Dia select x).ToList().FirstOrDefault();
                    busquedaindicadorDesarrollo = (from x in indicadorDesarrollo where x.FECHA.Day.ToString() == item.Dia select x).ToList().FirstOrDefault();
                    busquedaindicadorVaquillas = (from x in indicadorVaquillas where x.FECHA.Day.ToString() == item.Dia select x).ToList().FirstOrDefault();
                    busquedaindicadorSecas = (from x in indicadorSecas where x.FECHA.Day.ToString() == item.Dia select x).ToList().FirstOrDefault();
                    busquedaindicadorReto = (from x in indicadorReto where x.FECHA.Day.ToString() == item.Dia select x).ToList().FirstOrDefault();

                    item.Color_MH_Crecimiento = busquedaindicadorCrecimiento != null ? ColorDato(item.Crecimiento_Inventario, item.Crecimiento_MH, busquedaindicadorCrecimiento.MH) : "";
                    item.Color_MS_Crecimiento = busquedaindicadorCrecimiento != null ? ColorDato(item.Crecimiento_Inventario, item.Crecimiento_MS, busquedaindicadorCrecimiento.MS) : "";
                    item.Color_PorcentajeMs_Crecimiento = busquedaindicadorCrecimiento != null ? ColorDato(item.Crecimiento_Inventario, item.Crecimiento_PorcentajeMS, busquedaindicadorCrecimiento.PORCENTAJE_MS) : "";
                    item.Color_CostoMS_Crecimiento = busquedaindicadorCrecimiento != null ? ColorDato(item.Crecimiento_Inventario, item.Crecimiento_CostoMS, busquedaindicadorCrecimiento.KGMS) : "";
                    item.Color_Costo_Crecimiento = busquedaindicadorCrecimiento != null ? ColorDato(item.Crecimiento_Inventario, item.Crecimiento_Costo, busquedaindicadorCrecimiento.COSTO) : "";

                    item.Color_MH_Desarrollo = busquedaindicadorDesarrollo != null ? ColorDato(item.Desarrollo_Inventario, item.Desarrollo_MH, busquedaindicadorDesarrollo.MH) : "";
                    item.Color_MS_Desarrollo = busquedaindicadorDesarrollo != null ? ColorDato(item.Desarrollo_Inventario, item.Desarrollo_MS, busquedaindicadorDesarrollo.MS) : "";
                    item.Color_PorcentajeMs_Desarrollo = busquedaindicadorDesarrollo != null ? ColorDato(item.Desarrollo_Inventario, item.Desarrollo_PorcentajeMS, busquedaindicadorDesarrollo.PORCENTAJE_MS) : "";
                    item.Color_CostoMS_Desarrollo = busquedaindicadorDesarrollo != null ? ColorDato(item.Desarrollo_Inventario, item.Desarrollo_CostoMS, busquedaindicadorDesarrollo.KGMS) : "";
                    item.Color_Costo_Desarrollo = busquedaindicadorDesarrollo != null ? ColorDato(item.Desarrollo_Inventario, item.Desarrollo_Costo, busquedaindicadorDesarrollo.COSTO) : "";

                    item.Color_MH_Vaquillas = busquedaindicadorVaquillas != null ? ColorDato(item.Vaquillas_Inventario, item.Vaquillas_MH, busquedaindicadorVaquillas.MH) : "";
                    item.Color_MS_Vaquillas = busquedaindicadorVaquillas != null ? ColorDato(item.Vaquillas_Inventario, item.Vaquillas_MS, busquedaindicadorVaquillas.MS) : "";
                    item.Color_PorcentajeMs_Vaquillas = busquedaindicadorVaquillas != null ? ColorDato(item.Vaquillas_Inventario, item.Vaquillas_PorcentajeMS, busquedaindicadorVaquillas.PORCENTAJE_MS) : "";
                    item.Color_CostoMS_Vaquillas = busquedaindicadorVaquillas != null ? ColorDato(item.Vaquillas_Inventario, item.Vaquillas_CostoMS, busquedaindicadorVaquillas.KGMS) : "";
                    item.Color_Costo_Vaquillas = busquedaindicadorVaquillas != null ? ColorDato(item.Vaquillas_Inventario, item.Vaquillas_Costo, busquedaindicadorVaquillas.COSTO) : "";

                    item.Color_MH_Secas = busquedaindicadorSecas != null ? ColorDato(item.Secas_Inventario, item.Secas_MH, busquedaindicadorSecas.MH) : "";
                    item.Color_MS_Secas = busquedaindicadorSecas != null ? ColorDato(item.Secas_Inventario, item.Secas_MS, busquedaindicadorSecas.MS) : "";
                    item.Color_PorcentajeMs_Secas = busquedaindicadorSecas != null ? ColorDato(item.Secas_Inventario, item.Secas_PorcentajeMS, busquedaindicadorSecas.PORCENTAJE_MS) : "";
                    item.Color_CostoMS_Secas = busquedaindicadorSecas != null ? ColorDato(item.Secas_Inventario, item.Secas_CostoMS, busquedaindicadorSecas.KGMS) : "";
                    item.Color_Costo_Secas = busquedaindicadorSecas != null ? ColorDato(item.Secas_Inventario, item.Secas_Costo, busquedaindicadorSecas.COSTO) : "";

                    item.Color_MH_Reto = busquedaindicadorReto != null ? ColorDato(item.Reto_Inventario, item.Reto_MH, busquedaindicadorReto.MH) : "";
                    item.Color_MS_Reto = busquedaindicadorReto != null ? ColorDato(item.Reto_Inventario, item.Reto_MS, busquedaindicadorReto.MS) : "";
                    item.Color_PorcentajeMs_Reto = busquedaindicadorReto != null ? ColorDato(item.Reto_Inventario, item.Reto_PorcentajeMS, busquedaindicadorReto.PORCENTAJE_MS) : "";
                    item.Color_CostoMS_Reto = busquedaindicadorReto != null ? ColorDato(item.Reto_Inventario, item.Reto_CostoMS, busquedaindicadorReto.KGMS) : "";
                    item.Color_Costo_Reto = busquedaindicadorReto != null ? ColorDato(item.Reto_Inventario, item.Reto_Costo, busquedaindicadorReto.COSTO) : "";

                    item.Color_IXA = ColorUtilidad(item.IngresoxAnimal, utilidad.IXA, false);
                    item.Color_CXA = ColorUtilidad(item.CostoxAnimal, utilidad.CXA, true);
                    item.Color_PorcentajeCXA = ColorUtilidad(item.Porcentaje_CostoxAnimal, utilidad.PORCENTAJE_C, true);
                    item.Color_UXA = ColorUtilidad(item.UtilidadxAnimal, utilidad.UXA, false);
                    item.Color_PorcentajeUXA = ColorUtilidad(item.Porcentaje_UtilidadxAnimal, utilidad.PORCENTAJE_U, false);
                }
            }
            catch { }

        }

        public void CargarPrecioLeche(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            mensaje = string.Empty;
            bool esPrecioFacturado = EsPrecioLecheFacturado(fechaFin);

            precioLeche = esPrecioFacturado ? PrecioLecheFacturado(fechaInicio, fechaFin, ref mensaje) : PrecioLecheCalculado(rancho.Ran_ID, fechaInicio, fechaFin, ref mensaje);
        }

        #region hoja3
        public List<Hoja3> ReporteHoja3(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            List<Hoja3> response = new List<Hoja3>();
            mensaje = string.Empty;

            try
            {
                List<DateTime> fechasReporte = ListaFechasReporte(fechaFin);

                #region Obtener Datos
                List<CalostroYOrdeña> datosCalostro = CalostroOrdeño(fechaInicio, fechaFin);
                List<DatosVacas> datosJaulasVivas = DatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo <> 1 AND destino <> 2 AND (edovac = 1 OR edovac = 10) ", ref mensaje);
                List<DatosVacas> datosJaulasMuertas = DatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 1 AND (edovac = 1 OR edovac = 10) ", ref mensaje);

                List<DatosVacas> datosDesteteVivas = DatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo <> 1 AND destino <> 2 AND (edovac = 2 OR edovac = 7) ", ref mensaje);
                List<DatosVacas> datosDesteteMuertas = DatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 1 AND (edovac = 2 OR edovac = 7) ", ref mensaje);

                List<DatosVacas> datosVaquillasMuertas = DatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 1 AND destino = 1  AND (edovac = 3 OR edovac = 8 OR edovac = 9) ", ref mensaje);
                List<DatosVacas> datosVaquillasUrgencia = DatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 2 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ", ref mensaje);
                List<DatosVacas> datosVaquillasDelgadas = DatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 3 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ", ref mensaje);
                List<DatosVacas> datosVaquillasRegulares = DatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 4 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ", ref mensaje);
                List<DatosVacas> datosVaquillasGordas = DatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 5 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ", ref mensaje);
                List<DatosVacas> datosVaquillasOtros = DatosDesecho(fechaInicio, fechaFin, @"  AND vacvaq = 2  AND motivo > 5 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ", ref mensaje);

                List<DatosVacas> datosVacasMuertas = DatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo = 1 AND destino = 1 AND (edovac > 3) ", ref mensaje);
                List<DatosVacas> datosVacasUrgencia = DatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo = 2 AND destino = 1 AND (edovac > 3) ", ref mensaje);
                List<DatosVacas> datosVacasDelgadas = DatosDesecho(fechaInicio, fechaFin, @"  AND vacvaq = 1  AND motivo = 3 AND destino = 1 AND (edovac > 3) ", ref mensaje);
                List<DatosVacas> datosVacasRegulares = DatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo = 4 AND destino = 1 AND (edovac > 3) ", ref mensaje);
                List<DatosVacas> datosVacasGordas = DatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo = 5 AND destino = 1 AND (edovac > 3) ", ref mensaje);
                List<DatosVacas> datosVacasOtros = DatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo > 5 AND destino = 1 AND (edovac > 3) ", ref mensaje);

                List<DatosVacas> datosPartosVaquillasND = DatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2  AND (tipopar = 1 or tipopar = 3) ", ref mensaje);
                List<DatosVacas> datosPartosVaquillasAbortos = DatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2  AND (tipopar = 2) ", ref mensaje);
                List<DatosVacas> datosPartosVaquillasHembra = DatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2  AND criaviva = 1  AND criasexo = 1 ", ref mensaje);
                List<DatosVacas> datosPartosVaquillasMacho = DatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2  AND criaviva = 1  AND criasexo = 2 ", ref mensaje);

                List<DatosVacas> datosPartosVacasND = DatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 1  AND (tipopar = 1 or tipopar = 3) ", ref mensaje);
                List<DatosVacas> datosPartosVacasAbortos = DatosNacimiento(fechaInicio, fechaFin, @"  AND vacvaq = 1  AND (tipopar = 2) ", ref mensaje);
                List<DatosVacas> datosPartosVacasHembra = DatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 1  AND criaviva = 1  AND criasexo = 1 ", ref mensaje);
                List<DatosVacas> datosPartosVacasMacho = DatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 1  AND criaviva = 1  AND criasexo = 2 ", ref mensaje);

                List<DatosVacas> datosPartosSinCria = DatosNacimiento(fechaInicio, fechaFin, @" AND tipopar <> 2 AND criaviva = 2  AND criasexo = 3 ", ref mensaje);
                List<DatosVacas> datosPartosVacas = DatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 1 ", ref mensaje);
                List<DatosVacas> datosPartosVaquillas = DatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2 ", ref mensaje);

                List<DatosVacas> datosMuertasDia = DatosNacimientoMuertas(fechaInicio, fechaFin, @" AND criaviva = 2 AND (tipopar = 1 or tipopar = 3) AND criasexo < 3 AND dianoche = 1 ", ref mensaje);
                List<DatosVacas> datosMuertasNoche = DatosNacimientoMuertas(fechaInicio, fechaFin, @" AND criaviva = 2 AND (tipopar = 1 or tipopar = 3) AND criasexo < 3 AND dianoche = 2 ", ref mensaje);

                //List<DatosVacas> datosAbortosVacas = DatosAbortos(fechaInicio, fechaFin, 1, ref mensaje);
                //List<DatosVacas> datosAbortosVaquillas = DatosAbortos(fechaInicio, fechaFin, 2, ref mensaje);
                #endregion

                foreach (DateTime fecha in fechasReporte)
                {
                    Hoja3 item = new Hoja3();

                    item.Dia = fecha.Day.ToString();

                    #region obtener datos de la lista
                    CalostroYOrdeña busquedaCalostro = (from x in datosCalostro where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaJaulasVivas = (from x in datosJaulasVivas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaJaulasMuertas = (from x in datosJaulasMuertas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaDesteteVivas = (from x in datosDesteteVivas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaDesteteMuertas = (from x in datosDesteteMuertas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaVaquillasMuertas = (from x in datosVaquillasMuertas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaVaquillasUrgencia = (from x in datosVaquillasUrgencia where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaVaquillasDelgadas = (from x in datosVaquillasDelgadas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaVaquillasRegulares = (from x in datosVaquillasRegulares where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaVaquillasGordas = (from x in datosVaquillasGordas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaVaquillasOtros = (from x in datosVaquillasOtros where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaVacasMuertas = (from x in datosVacasMuertas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaVacasUrgencia = (from x in datosVacasUrgencia where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaVacasDelgadas = (from x in datosVacasDelgadas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaVacasRegulares = (from x in datosVacasRegulares where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaVacasGordas = (from x in datosVacasGordas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaVacasOtros = (from x in datosVacasOtros where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaPartosVaquillasND = (from x in datosPartosVaquillasND where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaPartosVaquillasAbortos = (from x in datosPartosVaquillasAbortos where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaPartosVaquillasHembra = (from x in datosPartosVaquillasHembra where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaPartosVaquillasMacho = (from x in datosPartosVaquillasMacho where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaPartosVacasND = (from x in datosPartosVacasND where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaPartosVacasAbortos = (from x in datosPartosVacasAbortos where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaPartosVacasHembra = (from x in datosPartosVacasHembra where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaPartosVacasMacho = (from x in datosPartosVacasMacho where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaPartosSinCria = (from x in datosPartosSinCria where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaPartosVacas = (from x in datosPartosVacas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaPartosVaquillas = (from x in datosPartosVaquillas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaMuertasDia = (from x in datosMuertasDia where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaMuertasNoche = (from x in datosMuertasNoche where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    //DatosVacas busquedaAbortosVacas = (from x in datosAbortosVacas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    //DatosVacas busquedaAbortosVaquillas = (from x in datosAbortosVaquillas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    #endregion

                    #region asignar datos
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
            catch { }


            return response;
        }

        public Hoja3 TotalReporteHoja3(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            Hoja3 response = new Hoja3() { Dia = "TOTAL" };
            mensaje = string.Empty;

            try
            {
                #region Obtener Datos
                CalostroYOrdeña totalCalostro = PromedioCalostro(fechaInicio, fechaFin);
                DatosVacas totalJaulasVivas = TotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo <> 1 AND destino <> 2 AND (edovac = 1 OR edovac = 10) ", ref mensaje);
                DatosVacas totalJaulasMuertas = TotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 1 AND (edovac = 1 OR edovac = 10) ", ref mensaje);

                DatosVacas totalDesteteVivas = TotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo <> 1 AND destino <> 2 AND (edovac = 2 OR edovac = 7) ", ref mensaje);
                DatosVacas totalDesteteMuertas = TotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 1 AND (edovac = 2 OR edovac = 7) ", ref mensaje);

                DatosVacas totalVaquillasMuertas = TotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 1 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ", ref mensaje);
                DatosVacas totalVaquillasUrgencia = TotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 2 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ", ref mensaje);
                DatosVacas totalVaquillasDelgadas = TotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 3 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ", ref mensaje);
                DatosVacas totalVaquillasRegulares = TotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 4 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ", ref mensaje);
                DatosVacas totalVaquillasGordas = TotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 2  AND motivo = 5 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ", ref mensaje);
                DatosVacas totalVaquillasOtros = TotalDatosDesecho(fechaInicio, fechaFin, @"  AND vacvaq = 2  AND motivo > 5 AND destino = 1 AND (edovac = 3 OR edovac = 8 OR edovac = 9) ", ref mensaje);

                DatosVacas totalVacasMuertas = TotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo = 1 AND destino = 1 AND (edovac > 3) ", ref mensaje);
                DatosVacas totalVacasUrgencia = TotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo = 2 AND destino = 1 AND (edovac > 3) ", ref mensaje);
                DatosVacas totalVacasDelgadas = TotalDatosDesecho(fechaInicio, fechaFin, @"  AND vacvaq = 1  AND motivo = 3 AND destino = 1 AND (edovac > 3) ", ref mensaje);
                DatosVacas totalVacasRegulares = TotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo = 4 AND destino = 1 AND (edovac > 3) ", ref mensaje);
                DatosVacas totalVacasGordas = TotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo = 5 AND destino = 1 AND (edovac > 3) ", ref mensaje);
                DatosVacas totalVacasOtros = TotalDatosDesecho(fechaInicio, fechaFin, @" AND vacvaq = 1  AND motivo > 5 AND destino = 1 AND (edovac > 3) ", ref mensaje);

                DatosVacas totalPartosVaquillasND = TotalDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2  AND (tipopar = 1 or tipopar = 3) ", ref mensaje);
                DatosVacas totalPartosVaquillasAbortos = TotalDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2  AND (tipopar = 2) ", ref mensaje);
                DatosVacas totalPartosVaquillasHembra = TotalDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2  AND criaviva = 1  AND criasexo = 1 ", ref mensaje);
                DatosVacas totalPartosVaquillasMacho = TotalDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2  AND criaviva = 1  AND criasexo = 2 ", ref mensaje);

                DatosVacas totalPartosVacasND = TotalDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 1  AND (tipopar = 1 or tipopar = 3) ", ref mensaje);
                DatosVacas totalPartosVacasAbortos = TotalDatosNacimiento(fechaInicio, fechaFin, @"  AND vacvaq = 1  AND (tipopar = 2) ", ref mensaje);
                DatosVacas totalPartosVacasHembra = TotalDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 1  AND criaviva = 1  AND criasexo = 1 ", ref mensaje);
                DatosVacas totalPartosVacasMacho = TotalDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 1  AND criaviva = 1  AND criasexo = 2 ", ref mensaje);

                DatosVacas totalPartosSinCria = TotalDatosNacimiento(fechaInicio, fechaFin, @" AND tipopar <> 2 AND criaviva = 2  AND criasexo = 3 ", ref mensaje);
                DatosVacas totalPartosVacas = TotalDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 1 ", ref mensaje);
                DatosVacas totalPartosVaquillas = TotalDatosNacimiento(fechaInicio, fechaFin, @" AND vacvaq = 2 ", ref mensaje);

                DatosVacas totalMuertasDia = TotalDatosNacimientoMuertas(fechaInicio, fechaFin, @" AND criaviva = 2 AND (tipopar = 1 or tipopar = 3) AND criasexo < 3 AND dianoche = 1 ", ref mensaje);
                DatosVacas totalMuertasNoche = TotalDatosNacimientoMuertas(fechaInicio, fechaFin, @" AND criaviva = 2 AND (tipopar = 1 or tipopar = 3) AND criasexo < 3 AND dianoche = 2 ", ref mensaje);

                DatosVacas totalAbortosVacas = TotalDatosAbortos(fechaInicio, fechaFin, 1, ref mensaje);
                DatosVacas totalAbortosVaquillas = TotalDatosAbortos(fechaInicio, fechaFin, 2, ref mensaje);
                #endregion

                #region asignar datos
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

        public Hoja3 DiferenciaReporteHoja3(Hoja3 total, Hoja3 totalAñoAnt, ref string mensaje)
        {
            Hoja3 response = new Hoja3() { Dia = "DIF #" };
            mensaje = string.Empty;

            try
            {
                #region asignar datos
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
            catch { }

            return response;
        }

        public Hoja3 PorcentajeDiferenciaReporteHoja3(Hoja3 diferencia, Hoja3 totalAñoAnt, ref string mensaje)
        {
            Hoja3 response = new Hoja3() { Dia = "DIF %" };
            mensaje = string.Empty;

            try
            {
                #region asignar datos
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
            catch { }

            return response;
        }

        public void QuitarCeros(List<Hoja3> reporte)
        {
            foreach (Hoja3 item in reporte)
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

                item.Diferencia_Calostro = item.Diferencia_Calostro == 0 ? null : item.Diferencia_Calostro;
                item.Porcentaje_Calostro = item.Porcentaje_Calostro == 0 ? null : item.Porcentaje_Calostro;
                item.Calidad_Calostro = item.Calidad_Calostro == 0 ? null : item.Calidad_Calostro;
            }
            catch { }
        }

        public List<Hoja3> EspaciosEnBlancoHoja3(int renglones)
        {
            List<Hoja3> response = new List<Hoja3>();
            int renglonesTotal = 31 - renglones;

            for (int i = 0; i < renglonesTotal; i++)
            {
                response.Add(new Hoja3());
            }

            return response;
        }
        #endregion

        #region hoja4
        public List<Hoja4> ReporteHoja4(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            List<Hoja4> response = new List<Hoja4>();
            mensaje = string.Empty;

            try
            {
                List<DateTime> fechasReporte = ListaFechasReporte(fechaFin);
                List<DatosVacas> datosUbre_MA = DatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 1 AND enfdetalle = 1 ", ref mensaje);
                List<DatosVacas> datosUbre_SL = DatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 1 AND enfdetalle = 2 ", ref mensaje);

                List<DatosVacas> datosMetabolicos_FL = DatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 5 AND enfdetalle = 1 ", ref mensaje);
                List<DatosVacas> datosMetabolicos_CET = DatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 5 AND enfdetalle = 2 ", ref mensaje);

                List<DatosVacas> datosLocomotores_BE = DatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 4 AND enfdetalle = 1 ", ref mensaje);
                List<DatosVacas> datosLocomotores_TRA = DatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 4 AND enfdetalle = 2 ", ref mensaje);
                List<DatosVacas> datosLocomotores_GA = DatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 4 AND enfdetalle = 3 ", ref mensaje);

                List<DatosVacas> datosDigestivos_AC = DatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 1 ", ref mensaje);
                List<DatosVacas> datosDigestivos_ES = DatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 2 ", ref mensaje);
                List<DatosVacas> datosDigestivos_DI = DatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 3 ", ref mensaje);
                List<DatosVacas> datosDigestivos_TI = DatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 4 ", ref mensaje);

                List<DatosVacas> datosReproductivos_RE = DatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 1 ", ref mensaje);
                List<DatosVacas> datosReproductivos_ME = DatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 2 ", ref mensaje);
                List<DatosVacas> datosReproductivos_PIO = DatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 3 ", ref mensaje);
                List<DatosVacas> datosReproductivos_QUI = DatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 4 ", ref mensaje);
                List<DatosVacas> datosReproductivos_CS = DatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 5 ", ref mensaje);

                List<DatosVacas> datosRespiratorios_Neu = DatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 7 AND enfdetalle = 1 ", ref mensaje);

                List<DatosVacas> datosBecerras_Neu = DatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 2  AND enfermedad = 8 AND enfdetalle = 1 ", ref mensaje);
                List<DatosVacas> datosBecerras_Fie = DatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 2  AND enfermedad = 8 AND enfdetalle = 2 ", ref mensaje);
                List<DatosVacas> datosBecerras_Di = DatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 2  AND enfermedad = 8 AND enfdetalle = 3 ", ref mensaje);
                List<DatosVacas> datosBecerras_Conj = DatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 2  AND enfermedad = 8 AND enfdetalle = 4 ", ref mensaje);
                List<ValInventario> datosValInventarios = DatosValInventario(fechaInicio, fechaFin, ref mensaje);
                List<DatosVacas> datosAbortosVacas = DatosAbortos(fechaInicio, fechaFin, 1, ref mensaje);
                List<DatosVacas> datosAbortosVaquillas = DatosAbortos(fechaInicio, fechaFin, 2, ref mensaje);

                foreach (DateTime fecha in fechasReporte)
                {
                    Hoja4 newItem = new Hoja4();
                    newItem.Dia = fecha.Day.ToString();

                    #region tomar datos por dia
                    DatosVacas busquedaUbre_MA = (from x in datosUbre_MA where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaUbre_SL = (from x in datosUbre_SL where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaMetabolicos_FL = (from x in datosMetabolicos_FL where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaMetabolicos_CET = (from x in datosMetabolicos_CET where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaLocomotores_BE = (from x in datosLocomotores_BE where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaLocomotores_TRA = (from x in datosLocomotores_TRA where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaLocomotores_GA = (from x in datosLocomotores_GA where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaDigestivos_AC = (from x in datosDigestivos_AC where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaDigestivos_ES = (from x in datosDigestivos_ES where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaDigestivos_DI = (from x in datosDigestivos_DI where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaDigestivos_TI = (from x in datosDigestivos_TI where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaReproductivos_RE = (from x in datosReproductivos_RE where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaReproductivos_ME = (from x in datosReproductivos_ME where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaReproductivos_PIO = (from x in datosReproductivos_PIO where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaReproductivos_QUI = (from x in datosReproductivos_QUI where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaReproductivos_CS = (from x in datosReproductivos_CS where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaRespiratorios_Neu = (from x in datosRespiratorios_Neu where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaBecerras_Neu = (from x in datosBecerras_Neu where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaBecerras_Fie = (from x in datosBecerras_Fie where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaBecerras_Di = (from x in datosBecerras_Di where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaBecerras_Conj = (from x in datosBecerras_Conj where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    ValInventario busquedaValInventarios = (from x in datosValInventarios where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaAbortosVacas = (from x in datosAbortosVacas where x.Fecha == fecha select x).ToList().FirstOrDefault();
                    DatosVacas busquedaAbortosVaquillas = (from x in datosAbortosVaquillas where x.Fecha == fecha select x).ToList().FirstOrDefault();
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
            catch (Exception ex) { mensaje = ex.Message; }

            return response;
        }

        public Hoja4 TotalHoja4(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            Hoja4 response = new Hoja4();
            response.Dia = "TOTAL";

            try
            {
                DatosVacas datosUbre_MA = TotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 1 AND enfdetalle = 1 ", ref mensaje);
                DatosVacas datosUbre_SL = TotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 1 AND enfdetalle = 2 ", ref mensaje);
                DatosVacas datosMetabolicos_FL = TotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 5 AND enfdetalle = 1 ", ref mensaje);
                DatosVacas datosMetabolicos_CET = TotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 5 AND enfdetalle = 2 ", ref mensaje);
                DatosVacas datosLocomotores_BE = TotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 4 AND enfdetalle = 1 ", ref mensaje);
                DatosVacas datosLocomotores_TRA = TotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 4 AND enfdetalle = 2 ", ref mensaje);
                DatosVacas datosLocomotores_GA = TotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 4 AND enfdetalle = 3 ", ref mensaje);
                DatosVacas datosDigestivos_AC = TotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 1 ", ref mensaje);
                DatosVacas datosDigestivos_ES = TotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 2 ", ref mensaje);
                DatosVacas datosDigestivos_DI = TotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 3 ", ref mensaje);
                DatosVacas datosDigestivos_TI = TotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 2 AND enfdetalle = 4 ", ref mensaje);
                DatosVacas datosReproductivos_RE = TotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 1 ", ref mensaje);
                DatosVacas datosReproductivos_ME = TotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 2 ", ref mensaje);
                DatosVacas datosReproductivos_PIO = TotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 3 ", ref mensaje);
                DatosVacas datosReproductivos_QUI = TotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 4 ", ref mensaje);
                DatosVacas datosReproductivos_CS = TotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 3 AND enfdetalle = 5 ", ref mensaje);
                DatosVacas datosRespiratorios_Neu = TotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 1  AND enfermedad = 7 AND enfdetalle = 1 ", ref mensaje);
                DatosVacas datosBecerras_Neu = TotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 2  AND enfermedad = 8 AND enfdetalle = 1 ", ref mensaje);
                DatosVacas datosBecerras_Fie = TotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 2  AND enfermedad = 8 AND enfdetalle = 2 ", ref mensaje);
                DatosVacas datosBecerras_Di = TotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 2  AND enfermedad = 8 AND enfdetalle = 3 ", ref mensaje);
                DatosVacas datosBecerras_Conj = TotalDatosSalud(fechaInicio, fechaFin, @" AND vacvaq = 2  AND enfermedad = 8 AND enfdetalle = 4 ", ref mensaje);
                ValInventario datosValInventarios = TotalDatosValInventario(fechaInicio, fechaFin, ref mensaje);
                DatosVacas datosAbortosVacas = TotalDatosAbortos(fechaInicio, fechaFin, 1, ref mensaje);
                DatosVacas datosAbortosVaquillas = TotalDatosAbortos(fechaInicio, fechaFin, 2, ref mensaje);

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

        public Hoja4 DiferenciHoja4(Hoja4 total, Hoja4 totalAñoAnt, ref string mensaje)
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
            catch { }

            return response;
        }

        public Hoja4 PorcentajeDiferenciHoja4(Hoja4 diferencia, Hoja4 totalAñoAnt, ref string mensaje)
        {
            Hoja4 response = new Hoja4();
            response.Dia = "DIF %";

            try
            {
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
            }
            catch { }

            return response;
        }

        public List<Hoja4> EspaciosEnBlancoHoja4(int renglones)
        {
            List<Hoja4> response = new List<Hoja4>();
            int renglonesTotal = 31 - renglones;

            for (int i = 0; i < renglonesTotal; i++)
            {
                response.Add(new Hoja4());
            }

            return response;
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

        public void CierreMesCorrecto(int ranId, int horaCorte, string urlWebService, DateTime fechaInicio, DateTime fechaFin, out bool validacionMedicina, out bool validacionAlimento)
        {
            bool validacionCCO = CentrosDeCostoValidos(ranId, urlWebService, fechaInicio, fechaFin);
            string mensaje = string.Empty;
            GTH.ValidacionesCierreMes(ranId, horaCorte, fechaInicio, fechaFin, out validacionAlimento, out validacionMedicina, ref mensaje);
            bool auxiliarMedicina = validacionCCO && validacionMedicina;
            validacionMedicina = auxiliarMedicina;
        }

        public DateTime FechaMaxima(ref string mensaje)
        {
            DateTime fecha = DateTime.Today;
            int juliana = ConvertToJulian(fecha);
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

            try
            {
                string query = @"SELECT MAX(FECHA) FROM VALINVENTARIO";

                db.Conectar();
                db.CrearComando(query, tipoComandoAccess.query);
                DataTable dt = db.EjecutarConsultaTabla();

                if (dt.Rows.Count > 0)
                    juliana = dt.Rows[0][0] != DBNull.Value ? Convert.ToInt32(dt.Rows[0][0]) : ConvertToJulian(fecha);

                fecha = ConvertJulianaToGregoriana(juliana);
            }
            catch (Exception ex) { mensaje = ex.Message; }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return fecha;
        }

        public DateTime FechaMinima(ref string mensaje)
        {
            DateTime fecha = DateTime.Today;
            int juliana = ConvertToJulian(fecha);
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

            try
            {
                string query = @"SELECT MIN(FECHA) FROM VALINVENTARIO";

                db.Conectar();
                db.CrearComando(query, tipoComandoAccess.query);
                DataTable dt = db.EjecutarConsultaTabla();

                if (dt.Rows.Count > 0)
                    juliana = dt.Rows[0][0] != DBNull.Value ? Convert.ToInt32(dt.Rows[0][0]) : ConvertToJulian(fecha);

                fecha = ConvertJulianaToGregoriana(juliana);
            }
            catch (Exception ex) { mensaje = ex.Message; }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return fecha;
        }
        #endregion

        #region ValoresConsistentes
        public void ValoresConsistentes(Rancho rancho, DateTime fecha)
        {
            try
            {
                ValConAlim valconAlim = new ValConAlim() { Ran_ID = rancho.Ran_ID };


                string mensaje = string.Empty;
                bool configuraDiario = ConfiguraDiario(ref mensaje);
                DateTime fechaFinal = !configuraDiario ? fecha.AddDays(-1) : fecha;
                bool evaluacionValConAlim = EvaluacionValConAlim(fechaFinal, ref mensaje);

                List<gth.ReportePeriodo> auxIngredientes = new List<gth.ReportePeriodo>();
                gth.ReportePeriodo auxSobrante = new gth.ReportePeriodo();

                gth.IndicadorReportePeriodo indicadoresProduccion = new gth.IndicadorReportePeriodo();
                gth.IndicadorReportePeriodo indicadoresSecas = new gth.IndicadorReportePeriodo();
                gth.IndicadorReportePeriodo indicadoresReto = new gth.IndicadorReportePeriodo();
                gth.IndicadorReportePeriodo indicadoresCrecimiento = new gth.IndicadorReportePeriodo();
                gth.IndicadorReportePeriodo indicadoresDesarrollo = new gth.IndicadorReportePeriodo();
                gth.IndicadorReportePeriodo indicadoresVaquillas = new gth.IndicadorReportePeriodo();

                if (evaluacionValConAlim)
                {
                    #region llamado dll Alimentacion
                    GTH.ReportePeriodo(rancho.Ran_ID.ToString(), rancho.TimeShiftTracker, "10,11,12,13", fechaFinal, fechaFinal, out auxIngredientes, out indicadoresProduccion, out auxSobrante);
                    GTH.ReportePeriodo(rancho.Ran_ID.ToString(), rancho.TimeShiftTracker, "21", fechaFinal, fechaFinal, out auxIngredientes, out indicadoresSecas, out auxSobrante);
                    GTH.ReportePeriodo(rancho.Ran_ID.ToString(), rancho.TimeShiftTracker, "22", fechaFinal, fechaFinal, out auxIngredientes, out indicadoresReto, out auxSobrante);
                    GTH.ReportePeriodo(rancho.Ran_ID.ToString(), rancho.TimeShiftTracker, "32", fechaFinal, fechaFinal, out auxIngredientes, out indicadoresCrecimiento, out auxSobrante);
                    GTH.ReportePeriodo(rancho.Ran_ID.ToString(), rancho.TimeShiftTracker, "33", fechaFinal, fechaFinal, out auxIngredientes, out indicadoresDesarrollo, out auxSobrante);
                    GTH.ReportePeriodo(rancho.Ran_ID.ToString(), rancho.TimeShiftTracker, "34", fechaFinal, fechaFinal, out auxIngredientes, out indicadoresVaquillas, out auxSobrante);
                    #endregion

                    #region Evaluacion de Ea Producción
                    if (indicadoresProduccion.ANIMELES == 0)
                    {
                        valconAlim.EAProd = true;
                        valconAlim.EAProdValor = -1;

                        valconAlim.MsProd = true;
                        valconAlim.MsProdValor = -1;
                    }
                    else
                    {
                        valconAlim.EAProd = (indicadoresProduccion.EAS >= 1.1 && indicadoresProduccion.EAS <= 1.8) ? true : false;
                        valconAlim.EAProdValor = indicadoresProduccion.EA;

                        valconAlim.MsProd = (indicadoresProduccion.MSS >= 18 && indicadoresProduccion.MSS <= 32) ? true : false;
                        valconAlim.MsProdValor = indicadoresProduccion.MSS;
                    }
                    #endregion

                    #region Evaluación de MS Destetadas Uno
                    if (indicadoresCrecimiento.ANIMELES == 0)
                    {
                        valconAlim.MsCrecimiento = true;
                        valconAlim.MsCrecimientosValor = -1;
                    }
                    else
                    {
                        valconAlim.MsCrecimiento = (indicadoresCrecimiento.MSS >= 2 && indicadoresCrecimiento.MSS <= 8) ? true : false;
                        valconAlim.MsCrecimientosValor = indicadoresCrecimiento.MSS;
                    }
                    #endregion

                    #region Evaluación de Ms Destete Dos
                    if (indicadoresDesarrollo.ANIMELES == 0)
                    {
                        valconAlim.MsDesarollo = true;
                        valconAlim.MsDesarolloValor = -1;
                    }
                    else
                    {
                        valconAlim.MsDesarollo = (indicadoresDesarrollo.MSS >= 6 && indicadoresDesarrollo.MSS <= 11) ? true : false;
                        valconAlim.MsDesarolloValor = indicadoresDesarrollo.MSS;
                    }
                    #endregion

                    #region Evaluación de Ms Vaquillas
                    if (indicadoresVaquillas.ANIMELES == 0)
                    {
                        valconAlim.MsVaquillas = true;
                        valconAlim.MsVaquillasValor = -1;
                    }
                    else
                    {
                        valconAlim.MsVaquillas = (indicadoresVaquillas.MSS >= 8 && indicadoresVaquillas.MSS <= 15) ? true : false;
                        valconAlim.MsVaquillasValor = indicadoresVaquillas.MSS;
                    }
                    #endregion

                    #region Evaluación de MS Secas
                    if (indicadoresSecas.ANIMELES == 0)
                    {
                        valconAlim.MsSecas = true;
                        valconAlim.MsSecasValor = -1;
                    }
                    else
                    {
                        valconAlim.MsSecas = (indicadoresSecas.MSS >= 10 && indicadoresSecas.MSS <= 19) ? true : false;
                        valconAlim.MsSecasValor = indicadoresSecas.MSS;
                    }
                    #endregion

                    #region Evaluación de MS Reto
                    if (indicadoresReto.ANIMELES == 0)
                    {
                        valconAlim.MsReto = true;
                        valconAlim.MsRetoValor = -1;
                    }
                    else
                    {
                        valconAlim.MsReto = (indicadoresReto.MSS >= 9 && indicadoresReto.MSS <= 19) ? true : false;
                        valconAlim.MsRetoValor = indicadoresReto.MSS;

                    }
                    #endregion

                    #region Evaluación de todas la etapas
                    if (valconAlim.EAProd && valconAlim.MsProd && valconAlim.MsCrecimiento && valconAlim.MsDesarollo && valconAlim.MsVaquillas && valconAlim.MsSecas && valconAlim.MsReto)
                    {
                        valconAlim.ValConsisitentes = true;
                    }
                    else
                    {
                        valconAlim.ValConsisitentes = false;
                    }
                    #endregion

                    InsertarValConAlim(valconAlim, fechaFinal, ref mensaje);
                }

                EliminarValConAlimPeriodo(ref mensaje);

                ValConAlim indicadorValConAlim = new ValConAlim();
                for (int i = 0; i < 36; i++)
                {
                    indicadorValConAlim = new ValConAlim();
                    InventarioAfiXDia invAFI = IventarioAFI(rancho.Ran_ID, fechaFinal.AddDays(-i), ref mensaje);
                    InventarioAfiXDia promedioInvAFI = PromedioIventarioAFI(rancho.Ran_ID, fechaFinal.AddDays(-i - 4), fechaFinal.AddDays(-i - 1), ref mensaje);
                    PromedioMProducAlimentacion infoValConAlimPerido = MProducAlim(fechaFinal.AddDays(-i - 4), fechaFinal.AddDays(-i - 1), ref mensaje);

                    #region asignar los valores ValConAlim
                    indicadorValConAlim.EAProdValor = Convert.ToDouble(infoValConAlimPerido.EA);
                    indicadorValConAlim.MsProdValor = Convert.ToDouble(infoValConAlimPerido.Produccion);
                    indicadorValConAlim.MsCrecimientosValor = Convert.ToDouble(infoValConAlimPerido.Crecimiento);
                    indicadorValConAlim.MsDesarolloValor = Convert.ToDouble(infoValConAlimPerido.Desarrollo);
                    indicadorValConAlim.MsVaquillasValor = Convert.ToDouble(infoValConAlimPerido.Vaquillas);
                    indicadorValConAlim.MsSecasValor = Convert.ToDouble(infoValConAlimPerido.Secas);
                    indicadorValConAlim.MsRetoValor = Convert.ToDouble(infoValConAlimPerido.Reto);
                    #endregion

                    if (invAFI != null)
                    {
                        #region evaluacion EA y MS PRODUCCION
                        if (invAFI.Ordeño == null || invAFI.Ordeño == 0)
                        {
                            indicadorValConAlim.EAProd = true;
                            indicadorValConAlim.EAProdValor = 0;

                            indicadorValConAlim.MsProd = true;
                            indicadorValConAlim.MsProdValor = 0;
                        }
                        else
                        {
                            indicadorValConAlim.EAProd = promedioInvAFI.Ordeño != null && promedioInvAFI.Ordeño > 0 ? (infoValConAlimPerido.EA >= 1.1M && infoValConAlimPerido.EA <= 1.8M) ? true : false : true;

                            indicadorValConAlim.MsProd = promedioInvAFI.Ordeño != null && promedioInvAFI.Ordeño > 0 ? (infoValConAlimPerido.Produccion >= 18 && infoValConAlimPerido.Produccion <= 32) ? true : false : true;
                        }
                        #endregion

                        #region evaluacion MS Crecimiento
                        if (invAFI.Crecimiento == null || invAFI.Crecimiento == 0)
                        {
                            indicadorValConAlim.MsCrecimiento = true;
                            indicadorValConAlim.MsCrecimientosValor = 0;
                        }
                        else
                        {
                            indicadorValConAlim.MsCrecimiento = promedioInvAFI.Crecimiento != null && promedioInvAFI.Crecimiento > 0 ? (infoValConAlimPerido.Crecimiento >= 2 && infoValConAlimPerido.Crecimiento <= 8) ? true : false : true;
                        }
                        #endregion

                        #region evaluacion MS Desarrollo
                        if (invAFI.Desarrollo == null || invAFI.Desarrollo == 0)
                        {
                            indicadorValConAlim.MsDesarollo = true;
                            indicadorValConAlim.MsDesarolloValor = 0;
                        }
                        else
                        {
                            indicadorValConAlim.MsDesarollo = promedioInvAFI.Desarrollo != null && promedioInvAFI.Desarrollo > 0 ? (infoValConAlimPerido.Desarrollo >= 6 && infoValConAlimPerido.Desarrollo <= 11) ? true : false : true;
                        }
                        #endregion

                        #region evaluacion MS Vaquillas
                        if (invAFI.Vaquillas == null || invAFI.Vaquillas == 0)
                        {
                            indicadorValConAlim.MsVaquillas = true;
                            indicadorValConAlim.MsVaquillasValor = 0;
                        }
                        else
                        {
                            indicadorValConAlim.MsVaquillas = promedioInvAFI.Vaquillas != null && promedioInvAFI.Vaquillas > 0 ? (infoValConAlimPerido.Vaquillas >= 8 && infoValConAlimPerido.Vaquillas <= 15) ? true : false : true;
                        }
                        #endregion

                        #region evaluacion MS Secas
                        if (invAFI.Vaquillas == null || invAFI.Vaquillas == 0)
                        {
                            indicadorValConAlim.MsSecas = true;
                            indicadorValConAlim.MsSecasValor = 0;
                        }
                        else
                        {
                            indicadorValConAlim.MsSecas = promedioInvAFI.Secas != null && promedioInvAFI.Secas > 0 ? (infoValConAlimPerido.Secas >= 10 && infoValConAlimPerido.Secas <= 19) ? true : false : true;

                        }
                        #endregion

                        #region evaluacion MS Reto
                        if (invAFI.Reto == null || invAFI.Reto == 0)
                        {
                            indicadorValConAlim.MsReto = true;
                            indicadorValConAlim.MsRetoValor = 0;
                        }
                        else
                        {
                            indicadorValConAlim.MsReto = promedioInvAFI.Reto != null && promedioInvAFI.Reto > 0 ? (infoValConAlimPerido.Reto >= 9 && infoValConAlimPerido.Reto <= 19) ? true : false : true;
                        }
                        #endregion

                        #region evaluacion todas las etpas
                        indicadorValConAlim.ValConsisitentes = indicadorValConAlim.EAProd && indicadorValConAlim.MsProd && indicadorValConAlim.MsCrecimiento && indicadorValConAlim.MsDesarollo && indicadorValConAlim.MsVaquillas && indicadorValConAlim.MsSecas && indicadorValConAlim.MsReto;

                        #endregion
                    }
                    else
                    {
                        indicadorValConAlim.EAProd = (infoValConAlimPerido.EA >= 1.1M && infoValConAlimPerido.EA <= 1.8M) ? true : false;
                        indicadorValConAlim.MsProd = (infoValConAlimPerido.Produccion >= 18 && infoValConAlimPerido.Produccion <= 32) ? true : false;
                        indicadorValConAlim.MsCrecimiento = (infoValConAlimPerido.Crecimiento >= 2 && infoValConAlimPerido.Crecimiento <= 8) ? true : false;
                        indicadorValConAlim.MsDesarollo = (infoValConAlimPerido.Desarrollo >= 6 && infoValConAlimPerido.Desarrollo <= 11) ? true : false;
                        indicadorValConAlim.MsVaquillas = (infoValConAlimPerido.Vaquillas >= 8 && infoValConAlimPerido.Vaquillas <= 15) ? true : false;
                        indicadorValConAlim.MsSecas = (infoValConAlimPerido.Secas >= 10 && infoValConAlimPerido.Secas <= 19) ? true : false;
                        indicadorValConAlim.MsReto = (infoValConAlimPerido.Reto >= 9 && infoValConAlimPerido.Reto <= 19) ? true : false;

                        indicadorValConAlim.ValConsisitentes = indicadorValConAlim.EAProd && indicadorValConAlim.MsProd && indicadorValConAlim.MsCrecimiento && indicadorValConAlim.MsDesarollo && indicadorValConAlim.MsVaquillas && indicadorValConAlim.MsSecas && indicadorValConAlim.MsReto;
                    }

                    InsertarValConAlimPeriodo(indicadorValConAlim, fechaFinal.AddDays(-i), ref mensaje);
                }


            }
            catch { }
        }
        #endregion


        #endregion

        #region metodos privados
        private List<DatosProduccion> DatosDeProduccion(DateTime fechaInicio, DateTime fechaFin)
        {
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
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
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@julianaI", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@julianaF", ConvertToJulian(fechaFin));
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
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }


            return response;

        }

        private DatosProduccion TotalDatosDeProduccion(DateTime fechaInicio, DateTime fechaFin)
        {
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
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
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@julianaI", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@julianaF", ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new DatosProduccion()
                {
                    Proteina = 0,
                    Urea = x["Urea1"] != DBNull.Value ? Convert.ToDecimal(x["Urea1"]) : 0,
                    Grasa = x["Grasa1"] != DBNull.Value ? Convert.ToDecimal(x["Grasa1"]) : 0,
                    CCS = x["CCS1"] != DBNull.Value ? Convert.ToDecimal(x["CCS1"]) : 0,
                    CTD = x["CTD1"] != DBNull.Value ? Convert.ToDecimal(x["CTD1"]) : 0
                }).ToList().FirstOrDefault();

            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }


            return response;

        }

        private int ConvertToJulian(DateTime Date)
        {
            TimeSpan ts = (Date - Convert.ToDateTime("01/01/1900"));
            int julianday = ts.Days + 2;
            return julianday;
        }

        private List<DateTime> ListaFechasReporte(DateTime fechaFin)
        {
            List<DateTime> listaFechas = new List<DateTime>();
            DateTime fechaInicio = new DateTime(fechaFin.Year, fechaFin.Month, 1);

            while (fechaInicio <= fechaFin)
            {
                listaFechas.Add(fechaInicio);
                fechaInicio = fechaInicio.AddDays(1);
            }

            return listaFechas;
        }

        private List<Mproduc> MediaProduccionNoIdReal(DateTime fechaInicio, DateTime fechaFin)
        {
            List<Mproduc> response = new List<Mproduc>();
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);

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
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@julianaI", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@julianaF", ConvertToJulian(fechaFin));
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
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private List<Mproduc> MediaProduccionNoId(DateTime fechaInicio, DateTime fechaFin)
        {
            List<Mproduc> response = new List<Mproduc>();
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);

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
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@julianaI", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@julianaF", ConvertToJulian(fechaFin));
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
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private List<Cribas> NivelCriba(DateTime fechaInicio, DateTime fechaFin)
        {
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            List<Cribas> response = new List<Cribas>();

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
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@julianaI", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@julianaF", ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new Cribas()
                {
                    Fecha = x["FechaG"] != DBNull.Value ? Convert.ToDateTime(x["FechaG"]) : new DateTime(),
                    Nivel1 = x["criba1"] != DBNull.Value ? Convert.ToDecimal(x["criba1"]) : 0,
                    Nivel2 = x["criba2"] != DBNull.Value ? Convert.ToDecimal(x["criba2"]) : 0,
                    Nivel3 = x["criba3"] != DBNull.Value ? Convert.ToDecimal(x["criba3"]) : 0,
                    Nivel4 = x["criba4"] != DBNull.Value ? Convert.ToDecimal(x["criba4"]) : 0

                }).ToList();

            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private List<CalostroYOrdeña> CalostroOrdeño(DateTime fechaInicio, DateTime fechaFin)
        {
            List<CalostroYOrdeña> response = new List<CalostroYOrdeña>();
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            int julianaInicio = ConvertToJulian(fechaInicio);
            int julianaFin = ConvertToJulian(fechaFin);

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
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@fechaInicio", julianaInicio);
                db.AsignarParametro("@fechaFin", julianaFin);
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
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }


            return response;
        }

        private List<Inventario> Inventario(DateTime fechaInicio, DateTime fechaFin)
        {
            List<Inventario> response = new List<Inventario>();
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);

            try
            {
                string query = @"SELECT  CDATE(Fecha) AS FechaG
                                        ,delord       AS Del
                                 FROM inventario
                                 WHERE fecha BETWEEN @fechaInicio AND @fechaFin";

                db.Conectar();
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@fechaInicio", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new Inventario()
                {
                    Fecha = x["FechaG"] != DBNull.Value ? Convert.ToDateTime(x["FechaG"]) : new DateTime(),
                    DelOrd = x["Del"] != DBNull.Value ? Convert.ToDecimal(x["Del"]) : 0
                }).ToList();

            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private PromedioHoja1 PromedioHoja1NoIdReal(DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            PromedioHoja1 response = new PromedioHoja1();
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

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
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@inicio", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fin", ConvertToJulian(fechaFin));
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
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }


            return response;
        }

        private PromedioHoja1 PromedioHoja1NoId(DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            PromedioHoja1 response = new PromedioHoja1();
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

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
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@inicio", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fin", ConvertToJulian(fechaFin));
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
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }


            return response;
        }

        private CalostroYOrdeña PromedioCalostro(DateTime fechaInicio, DateTime fechaFin)
        {
            CalostroYOrdeña response = new CalostroYOrdeña() { Fecha = fechaFin };
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);

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
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@fechaInicio", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", ConvertToJulian(fechaFin));
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
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private List<BacteriologiaLeche> BacteriologiaLeche(DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            List<BacteriologiaLeche> response = new List<BacteriologiaLeche>();
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

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
                db.CrearComando(query, tipoComandoAccess.query);
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new BacteriologiaLeche()
                {
                    FECHA = x["FECHA"] != DBNull.Value ? Convert.ToDateTime(x["FECHA"]) : new DateTime(1, 1, 1),
                    BACTEROLOGIA = x["PRECIO_BACTERIOLOGIA"] != DBNull.Value ? Convert.ToDecimal(x["PRECIO_BACTERIOLOGIA"]) : 0
                }).ToList();
            }
            catch (Exception ex) { mensaje = ex.Message; }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

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



                //if (valor != null)
                //{
                //    if (valor != 0)
                //    {
                //        switch (valor)
                //        {
                //            //Color Verde
                //            case decimal n when ((n >= (promedio * 0.95M) && n <= (promedio * 1.05M))):
                //                color = "#DEEDD3";
                //                break;
                //            //Color Blanco
                //            case decimal n when (n >= (promedio * 0.9M) && n < (promedio * 0.95M)):
                //                color = "#F2F2F2";
                //                break;
                //            case decimal n when (n > (promedio * 1.05M) && n <= (promedio * 1.1M)):
                //                color = "#F2F2F2";
                //                break;
                //            //Color Amarillo
                //            case decimal n when (n >= (promedio * 0.85M) && n < (promedio * 0.9M)):
                //                color = "#FFF5D9";
                //                break;
                //            case decimal n when (n > (promedio * 1.1M) && n <= (promedio * 1.15M)):
                //                color = "#FFF5D9";
                //                break;
                //            //Color Rojo
                //            case decimal n when (n < (promedio * 0.85M) || n > (promedio * 1.15M)):
                //                color = "#FFC9C9";
                //                break;
                //        }
                //    }
                //    else if (inventario != 0)
                //    {
                //        color = "#FFC9C9";
                //    }
                //}
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

        private List<InventarioAfiXDia> InventariosAFI(DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            List<InventarioAfiXDia> response = new List<InventarioAfiXDia>();
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

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
                                        ,(jaulas + destetadas + destetadas2 + vaquillas + vacassecas + vacasordeña + vqreto + vcreto) AS InventarioTotal
                                FROM INVENTARIOAFIXDIA
                                WHERE fecha BETWEEN @fechaInicio AND @fechaFin
                                ORDER BY 1";

                db.Conectar();
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@fechaInicio", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", ConvertToJulian(fechaFin));
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
                    InventarioTotal = x["InventarioTotal"] != DBNull.Value ? Convert.ToDecimal(x["InventarioTotal"]) : 0
                }).ToList();
            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }


            return response;
        }

        private InventarioAfiXDia PromedioInventariosAFI(DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            InventarioAfiXDia response = new InventarioAfiXDia();
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
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
                                        ,Round(SUM((jaulas + destetadas + destetadas2 + vaquillas + vacassecas + vacasordeña + vqreto + vcreto))/(@fechaFin - @fechaInicio + 1),0) AS InventarioTotal
                                FROM INVENTARIOAFIXDIA
                                WHERE fecha BETWEEN @fechaInicio AND @fechaFin";
                query = query.Replace("@fechaInicio", ConvertToJulian(fechaInicio).ToString()).Replace("@fechaFin", ConvertToJulian(fechaFin).ToString());

                db.Conectar();
                db.CrearComando(query, tipoComandoAccess.query);
                //db.AsignarParametro("@fechaInicio", ConvertToJulian(fechaInicio));
                //db.AsignarParametro("@fechaFin", ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new InventarioAfiXDia()
                {
                    Ordeño = x["PRODUCCION"] != DBNull.Value ? Convert.ToDecimal(x["PRODUCCION"]) : 0,
                    Secas = x["SECAS"] != DBNull.Value ? Convert.ToDecimal(x["SECAS"]) : 0,
                    Reto = x["RETO"] != DBNull.Value ? Convert.ToDecimal(x["RETO"]) : 0,
                    Jaulas = x["JAULAS_"] != DBNull.Value ? Convert.ToDecimal(x["JAULAS_"]) : 0,
                    Crecimiento = x["CRECIMIENTO"] != DBNull.Value ? Convert.ToDecimal(x["CRECIMIENTO"]) : 0,
                    Desarrollo = x["DESARROLLO"] != DBNull.Value ? Convert.ToDecimal(x["DESARROLLO"]) : 0,
                    Vaquillas = x["VAQUILLAS_"] != DBNull.Value ? Convert.ToDecimal(x["VAQUILLAS_"]) : 0,
                    InventarioTotal = x["InventarioTotal"] != DBNull.Value ? Convert.ToDecimal(x["InventarioTotal"]) : 0
                }).ToList().FirstOrDefault();
            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }


            return response;
        }

        private List<Mproduc2> MediaProduccion2(DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            List<Mproduc2> response = new List<Mproduc2>();
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

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
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@fechaInicio", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", ConvertToJulian(fechaFin));
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
                mensaje = ex.Message;
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

        private decimal PrecioLecheFacturado(DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            decimal precio = 0;
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

            try
            {
                string query = @"SELECT AVG(PRECIOLECHE) AS PRECIO_LECHE
                                 FROM MPRODUC m
                                 WHERE m.FECHA BETWEEN @fechaInicio AND @fechaFin";

                db.Conectar();
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@fechaInicio", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", ConvertToJulian(fechaFin));
                DataTable dt = db.EjecutarConsultaTabla();

                if (dt.Rows.Count > 0)
                    precio = dt.Rows[0]["PRECIO_LECHE"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["PRECIO_LECHE"]) : 0;
            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return precio;
        }

        private decimal PrecioLecheCalculado(int ranId, DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            decimal precio = 0;
            ModeloDatosSQL db = new ModeloDatosSQL(conexionSQL);
            mensaje = string.Empty;

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
	                                SELECT  Ran_id
	                                       ,MAX(Pre_fecha) AS Fecha
	                                FROM PRECIO_LECHE_HISTORICO
	                                WHERE Ran_id = @rancho
                                        AND CONVERT(INT,FORMAT(Pre_Fecha, 'yyyyMM')) = @fecha
	                                GROUP BY  Ran_id
                                )
                                SELECT  plh.Precio_leche AS Precio_Leche
                                FROM PRECIO_LECHE_HISTORICO plh
                                INNER JOIN FechaPrecioLeche fpl
                                ON plh.Ran_id = fpl.Ran_id AND plh.Pre_Fecha = fpl.Fecha";

                    db.CrearComando(query, tipoComandoSQL.query);
                    db.AsignarParametro("@rancho", ranId);
                    db.AsignarParametro("@fecha", Convert.ToInt32(fechaFin.ToString("yyyyMM")));
                }
                else
                {
                    query = @"WITH FechaPrecioLeche AS
                                (
	                                SELECT  Ran_id
	                                       ,MAX(Pre_fecha) AS Fecha
	                                FROM PRECIO_LECHE_HISTORICO
	                                WHERE Ran_id = @rancho
	                                GROUP BY  Ran_id
                                )
                                SELECT  plh.Precio_leche AS Precio_Leche
                                FROM PRECIO_LECHE_HISTORICO plh
                                INNER JOIN FechaPrecioLeche fpl
                                ON plh.Ran_id = fpl.Ran_id AND plh.Pre_Fecha = fpl.Fecha";
                    db.CrearComando(query, tipoComandoSQL.query);
                    db.AsignarParametro("@rancho", ranId);
                }

                DataTable dt = db.EjecutarConsultaTabla();

                if (dt.Rows.Count > 0)
                    precio = dt.Rows[0]["Precio_Leche"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["Precio_Leche"]) : 0;

            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return precio;
        }

        private decimal LecheProducida(DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            decimal lecheProducida = 0;
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

            try
            {
                string query = @"SELECT IIF( SUM(IIF(m.LECPROD > 0, 1,0)) > 0, CLng(SUM(m.LECPROD) /  SUM(IIF(m.LECPROD > 0, 1,0))), 0) AS LecheProd
                                 FROM MPRODUC m
                                 WHERE m.FECHA BETWEEN @fechaInicio AND @fechaFin";

                db.Conectar();
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@fechaInicio", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", ConvertToJulian(fechaFin));
                DataTable dt = db.EjecutarConsultaTabla();

                if (dt.Rows.Count > 0)
                    lecheProducida = dt.Rows[0]["LecheProd"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["LecheProd"]) : 0;

            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return lecheProducida;
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

        private List<DatosVacas> DatosDesecho(DateTime fechaInicio, DateTime fechaFin, string condicion, ref string mensaje)
        {
            List<DatosVacas> response = new List<DatosVacas>();
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);

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
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@fechaInicio", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new DatosVacas()
                {
                    Fecha = x["FechaG"] != DBNull.Value ? Convert.ToDateTime(x["FechaG"]) : new DateTime(),
                    Vacas = x["Vacas"] != DBNull.Value ? Convert.ToDecimal(x["Vacas"]) : 0
                }).ToList();
            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private List<DatosVacas> DatosNacimiento(DateTime fechaInicio, DateTime fechaFin, string condicion, ref string mensaje)
        {
            List<DatosVacas> response = new List<DatosVacas>();
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

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
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@fechaInicio", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new DatosVacas()
                {
                    Fecha = x["FechaG"] != DBNull.Value ? Convert.ToDateTime(x["FechaG"]) : new DateTime(),
                    Vacas = x["Vacas"] != DBNull.Value ? Convert.ToDecimal(x["Vacas"]) : 0
                }).ToList();
            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private List<DatosVacas> DatosNacimientoMuertas(DateTime fechaInicio, DateTime fechaFin, string condicion, ref string mensaje)
        {
            List<DatosVacas> response = new List<DatosVacas>();
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

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
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@fechaInicio", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new DatosVacas()
                {
                    Fecha = x["FechaG"] != DBNull.Value ? Convert.ToDateTime(x["FechaG"]) : new DateTime(),
                    Vacas = x["Vacas"] != DBNull.Value ? Convert.ToDecimal(x["Vacas"]) : 0
                }).ToList();

            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private List<DatosVacas> DatosAbortos(DateTime fechaInicio, DateTime fechaFin, int condicion, ref string mensaje)
        {
            List<DatosVacas> response = new List<DatosVacas>();
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

            try
            {
                string query = @"SELECT CDATE(FECHA) AS FechaG
                                        , COUNT(*) AS Vacas
                                 FROM Dabortos 
                                 WHERE FECHA BETWEEN @fechaInicio AND @fechaFin
                                        AND vacvaq = @tipo
                                 GROUP BY FECHA  ORDER BY FECHA";

                db.Conectar();
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@fechaInicio", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", ConvertToJulian(fechaFin));
                db.AsignarParametro("@tipo", condicion);
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new DatosVacas()
                {
                    Fecha = x["FechaG"] != DBNull.Value ? Convert.ToDateTime(x["FechaG"]) : new DateTime(),
                    Vacas = x["Vacas"] != DBNull.Value ? Convert.ToDecimal(x["Vacas"]) : 0
                }).ToList();
            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private DatosVacas TotalDatosDesecho(DateTime fechaInicio, DateTime fechaFin, string condicion, ref string mensaje)
        {
            DatosVacas response = new DatosVacas();
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);

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
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@fechaInicio", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", ConvertToJulian(fechaFin));
                DataTable dt = db.EjecutarConsultaTabla();

                if (dt.Rows.Count > 0)
                    response.Vacas = dt.Rows[0]["Vacas_"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["Vacas_"]) : 0;
            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private DatosVacas TotalDatosNacimiento(DateTime fechaInicio, DateTime fechaFin, string condicion, ref string mensaje)
        {
            DatosVacas response = new DatosVacas();
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

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
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@fechaInicio", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", ConvertToJulian(fechaFin));
                DataTable dt = db.EjecutarConsultaTabla();

                if (dt.Rows.Count > 0)
                    response.Vacas = dt.Rows[0]["Vacas_"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["Vacas_"]) : 0;
            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private DatosVacas TotalDatosNacimientoMuertas(DateTime fechaInicio, DateTime fechaFin, string condicion, ref string mensaje)
        {
            DatosVacas response = new DatosVacas();
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

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
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@fechaInicio", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", ConvertToJulian(fechaFin));
                DataTable dt = db.EjecutarConsultaTabla();

                if (dt.Rows.Count > 0)
                    response.Vacas = dt.Rows[0]["Vacas_"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["Vacas_"]) : 0;

            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private DatosVacas TotalDatosAbortos(DateTime fechaInicio, DateTime fechaFin, int condicion, ref string mensaje)
        {
            DatosVacas response = new DatosVacas();
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

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
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@fechaInicio", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", ConvertToJulian(fechaFin));
                db.AsignarParametro("@tipo", condicion);
                DataTable dt = db.EjecutarConsultaTabla();

                if (dt.Rows.Count > 0)
                    response.Vacas = dt.Rows[0]["Vacas_"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["Vacas_"]) : 0;
            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private List<DatosVacas> DatosSalud(DateTime fechaInicio, DateTime fechaFin, string condicion, ref string mensaje)
        {
            List<DatosVacas> response = new List<DatosVacas>();
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

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
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@fechaInicio", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", ConvertToJulian(fechaFin));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new DatosVacas()
                {
                    Fecha = x["FechaG"] != DBNull.Value ? Convert.ToDateTime(x["FechaG"]) : new DateTime(),
                    Vacas = x["Vacas"] != DBNull.Value ? Convert.ToDecimal(x["Vacas"]) : 0
                }).ToList();

            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private DatosVacas TotalDatosSalud(DateTime fechaInicio, DateTime fechaFin, string condicion, ref string mensaje)
        {
            DatosVacas response = new DatosVacas();
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

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
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@fechaInicio", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", ConvertToJulian(fechaFin));
                DataTable dt = db.EjecutarConsultaTabla();

                if (dt.Rows.Count > 0)
                    response.Vacas = dt.Rows[0]["Vacas_"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["Vacas_"]) : 0;

            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private List<ValInventario> DatosValInventario(DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            List<ValInventario> response = new List<ValInventario>();
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

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
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@fechaInicio", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", ConvertToJulian(fechaFin));
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
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private ValInventario TotalDatosValInventario(DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            ValInventario response = new ValInventario();
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

            try
            {
                string query = @"SELECT  SUM(preñvc)       AS VacasPreñadas
                                        ,SUM(vaciavc)      AS VacasVacias
                                        ,SUM(preñvq)       AS VaquillasPreñadas
                                        ,SUM(vaciavq)      AS VaquillasVacias
                                FROM VALINVENTARIO
                                WHERE FECHA BETWEEN @fechaInicio AND @fechaFin";

                db.Conectar();
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@fechaInicio", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", ConvertToJulian(fechaFin));
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
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private bool CentrosDeCostoValidos(int ranId, string urlWebService, DateTime fechaInicio, DateTime fechaFin)
        {
            bool valido = true;

            try
            {
                List<ConfiguracionRancho> listaConfiguracion = ConfiguracionEstablo(ranId);
                List<string> listaSuscursal = ListaSuscursal(ranId);
                List<CentroDeCosto> listaCentrosDeCosto = ListaCentrosDeCosto(ranId, fechaInicio, fechaFin);

                string erp = listaConfiguracion[0].Erp_id;
                string susculsal = listaSuscursal[0];
                int dia = 0;
                int empresa = listaConfiguracion[0].Emp_prorrateo == 0 ? listaConfiguracion[0].Emp_id : listaConfiguracion[0].Emp_prorrateo;
                string url = urlWebService.Replace("@erp", listaConfiguracion[0].Erp_id + "erp");

                ght001766 sie = new ght001766(url, "", "", "");
                wCCO66DataTable centroCosto = new wCCO66DataTable();

                List<string> listaCentrosNoValidos = new List<string>();
                try
                {
                    sie.ght001766q(empresa, out centroCosto);

                    List<CCOGHT> listaDll = centroCosto.AsEnumerable().Select(x => new CCOGHT()
                    {
                        CCO_Clave = x["CcoClave"] != DBNull.Value ? x["CcoClave"].ToString() : string.Empty,
                        CCO_Nombre = x["CcoNombre"] != DBNull.Value ? x["CcoNombre"].ToString() : string.Empty
                    }).ToList();


                    foreach (CentroDeCosto item in listaCentrosDeCosto)
                    {
                        List<CCOGHT> busquedaCentroCostos = (from x in listaDll where x.CCO_Clave == item.CCO select x).ToList();

                        if (busquedaCentroCostos.Count == 0)
                        {
                            listaCentrosNoValidos.Add(item.CCO);
                        }
                    }

                    valido = listaCentrosNoValidos.Count == 0;
                }
                catch { }
                finally
                {
                    sie.Dispose();
                }
            }
            catch
            {

            }

            return valido;
        }

        private List<ConfiguracionRancho> ConfiguracionEstablo(int ranId)
        {
            List<ConfiguracionRancho> lista = new List<ConfiguracionRancho>();
            ModeloDatosSQL bd = new ModeloDatosSQL(conexionSQL);

            try
            {
                string query = @"";
                if (ranId == 25)
                {
                    query = @"Select erp_id, emp_id, track_id, ran_sie , emp_prorrateo
						FROM DBSIO.dbo.configuracion
						where ran_id IN( 25,29)";
                }
                else
                {
                    query = @"Select erp_id, emp_id, track_id, ran_sie ,[emp_prorrateo]
						FROM DBSIO.dbo.configuracion
						where ran_id IN(@rancho)";

                    query = query.Replace("@rancho", ranId.ToString());
                }

                bd.Conectar();
                bd.CrearComando(query, tipoComandoSQL.query);
                lista = bd.EjecutarConsultaTabla().AsEnumerable().Select(x => new ConfiguracionRancho()
                {
                    Erp_id = x["erp_id"] != DBNull.Value ? x["erp_id"].ToString() : string.Empty,
                    Emp_id = x["emp_id"] != DBNull.Value ? Convert.ToInt32(x["emp_id"]) : 0,
                    Track_id = x["track_id"] != DBNull.Value ? Convert.ToInt32(x["track_id"]) : 0,
                    Ran_sie = x["ran_sie"] != DBNull.Value ? Convert.ToInt32(x["ran_sie"]) : 0,
                    Emp_prorrateo = x["emp_prorrateo"] != DBNull.Value ? Convert.ToInt32(x["emp_prorrateo"]) : 0
                }).ToList();
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            finally
            {
                if (bd.isConnected)
                    bd.Desconectar();
            }

            return lista;
        }

        private List<string> ListaSuscursal(int ranId)
        {
            List<string> lista = new List<string>();
            ModeloDatosSQL bd = new ModeloDatosSQL(conexionSQL);

            try
            {
                string query = @"";

                if (ranId == 25)
                {
                    query = @"SELECT suc_id 
							  FROM DBSIO.dbo.RANCHO_SUCURSAL 
							  WHERE ran_id IN(25,29)";
                }
                else
                {
                    query = @"SELECT suc_id 
							  FROM DBSIO.dbo.RANCHO_SUCURSAL 
							  WHERE ran_id IN(@rancho)";
                    query = query.Replace("@rancho", ranId.ToString());
                }

                bd.Conectar();
                bd.CrearComando(query, tipoComandoSQL.query);
                DataTable dt = bd.EjecutarConsultaTabla();

                foreach (DataRow row in dt.Rows)
                    lista.Add(row["suc_id"].ToString());

            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            finally
            {
                if (bd.isConnected)
                    bd.Desconectar();
            }

            return lista;
        }

        private List<CentroDeCosto> ListaCentrosDeCosto(int ranId, DateTime fechaInicio, DateTime fechaFin)
        {
            List<CentroDeCosto> lista = new List<CentroDeCosto>();
            ModeloDatosSQL bd = new ModeloDatosSQL(conexionSQL);

            try
            {
                string query = @"";

                if (ranId == 25)
                {
                    query = @"SELECT DISTINCT Convert(nvarchar(50),RTRIM([CCO])) as CCO
							  FROM [DBALIMENTO].[dbo].[MEDICINA_ACTUAL]
							  WHERE 
							  Clave_Medicamento IS NOT NULL AND
							  Fecha >= @fechaInicio
							  AND Fecha <= @fechaFin
							  AND CCO <> 'POR ASIGNAR'
							  AND   (Clave_Medicamento LIKE 'MEBA[0-9][0-9][0-9][0-9]%'
							  OR   Clave_Medicamento LIKE 'RPHO[0-9][0-9][0-9][0-9]%'
							  OR   Clave_Medicamento LIKE 'VAVG[0-9][0-9][0-9][0-9]%')
							  UNION 
							  SELECT Convert(nvarchar(50),[cco_vacas])
	      					  FROM   [DBSIO].[dbo].[CCO_VAC_VAQ]
							  WHERE ran_id IN (25,29)
							  UNION 
							  SELECT Convert(nvarchar(50),[cco_vaquillas])
	      					  FROM   [DBSIO].[dbo].[CCO_VAC_VAQ]
							  WHERE ran_id IN (25,29)
							  ORDER BY CCO";
                }
                else
                {
                    query = @" SELECT DISTINCT Convert(nvarchar(50), RTRIM([CCO])) as CCO
							  FROM[DBALIMENTO].[dbo].[MEDICINA_ACTUAL]
							  WHERE
							  Clave_Medicamento IS NOT NULL AND
							  Fecha >= @fechaInicio
							  AND Fecha <= @fechaFin
							  AND CCO<> 'POR ASIGNAR'
							  AND(Clave_Medicamento LIKE 'MEBA[0-9][0-9][0-9][0-9]%'
							  OR   Clave_Medicamento LIKE 'RPHO[0-9][0-9][0-9][0-9]%'
							  OR   Clave_Medicamento LIKE 'VAVG[0-9][0-9][0-9][0-9]%')
							  UNION
							  SELECT Convert(nvarchar(50),[cco_vacas])
								FROM[DBSIO].[dbo].[CCO_VAC_VAQ]
							  WHERE ran_id = @ranId
							  UNION
							  SELECT Convert(nvarchar(50),[cco_vaquillas])
								FROM[DBSIO].[dbo].[CCO_VAC_VAQ]
							  WHERE ran_id = @ranId
							  ORDER BY CCO";
                    query = query.Replace("@ranId", ranId.ToString());
                }

                bd.Conectar();
                bd.CrearComando(query, tipoComandoSQL.query);
                bd.AsignarParametro("@fechaInicio", fechaInicio);
                bd.AsignarParametro("@fechaFin", fechaFin);
                lista = bd.EjecutarConsultaTabla().AsEnumerable().Select(x => new CentroDeCosto()
                {
                    CCO = x["CCO"] != DBNull.Value ? x["CCO"].ToString() : string.Empty
                }).ToList();


            }
            catch { }


            return lista;
        }

        private DateTime ConvertJulianaToGregoriana(int juliana)
        {
            DateTime date = new DateTime(1900, 1, 1);
            date = date.AddDays(juliana).AddDays(-2);
            return date;
        }

        private bool ConfiguraDiario(ref string mensaje)
        {
            bool configura = true;
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

            try
            {
                string query = "Select Configura from ConfiguraDiario";
                db.Conectar();
                db.CrearComando(query, tipoComandoAccess.query);
                DataTable dt = db.EjecutarConsultaTabla();

                if (dt.Rows.Count > 0)
                    configura = dt.Rows[0]["Configura"] != DBNull.Value ? Convert.ToBoolean(dt.Rows[0]["Configura"]) : true;

            }
            catch (Exception ex) { mensaje = ex.Message; }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return configura;
        }

        private bool EvaluacionValConAlim(DateTime fecha, ref string mensaje)
        {
            bool response = true;
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

            try
            {
                string query = "Select * from ValConAlim WHERE fecha = #@FechaDia#";
                query = query.Replace("@FechaDia", fecha.ToString("yyyy/MM/dd"));

                db.Conectar();
                db.CrearComando(query, tipoComandoAccess.query);
                DataTable dt = db.EjecutarConsultaTabla();

                response = dt.Rows.Count == 0;
            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        private void InsertarValConAlim(ValConAlim valconAlim, DateTime fecha, ref string mensaje)
        {
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

            try
            {
                string query = @"Insert into ValConAlim values (@Id, #@Fecha#, @EAProd, @MsProd, @MsCrecimiento, @MsDesarollo, @MsVaquillas, @MsSecas, @MsReto, @ValConsisitentes, @ValEAProd, @ValMsProd, @ValMsCrecimiento, @ValMsDesarollo, @ValMsVaquillas, @ValMsSecas, @ValMsReto)";
                query = query.Replace("@Id", valconAlim.Ran_ID.ToString())
                            .Replace("@Fecha", fecha.ToString("yyyy/MM/dd"))
                            .Replace("@EAProd", valconAlim.EAProd.ToString())
                            .Replace("@MsProd", valconAlim.MsProd.ToString())
                            .Replace("@MsCrecimiento", valconAlim.MsCrecimiento.ToString())
                            .Replace("@MsDesarollo", valconAlim.MsDesarollo.ToString())
                            .Replace("@MsVaquillas", valconAlim.MsVaquillas.ToString())
                            .Replace("@MsSecas", valconAlim.MsSecas.ToString())
                            .Replace("@MsReto", valconAlim.MsReto.ToString())
                            .Replace("@ValConsisitentes", valconAlim.ValConsisitentes.ToString())
                            .Replace("@ValEAProd", valconAlim.EAProdValor.ToString())
                            .Replace("@ValMsProd", valconAlim.MsProdValor.ToString())
                            .Replace("@ValMsCrecimiento", valconAlim.MsCrecimientosValor.ToString())
                            .Replace("@ValMsDesarollo", valconAlim.MsDesarolloValor.ToString())
                            .Replace("@ValMsVaquillas", valconAlim.MsVaquillasValor.ToString())
                            .Replace("@ValMsSecas", valconAlim.MsSecasValor.ToString())
                            .Replace("@ValMsReto", valconAlim.MsRetoValor.ToString());

                db.Conectar();
                db.CrearComando(query, tipoComandoAccess.query);
                int insert = db.EjecutarComando();
            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }
        }

        private void EliminarValConAlimPeriodo(ref string mensaje)
        {
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

            try
            {
                string query = @"DELETE FROM ValConAlimPeriodo";
                db.Conectar();
                db.CrearComando(query, tipoComandoAccess.query);
                int delete = db.EjecutarComando();
                Console.WriteLine("{0} Elementos Eliminados", delete);
            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }
        }

        private void InsertarValConAlimPeriodo(ValConAlim valconAlim, DateTime fecha, ref string mensaje)
        {
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

            try
            {
                string query = "Insert into ValConAlimPeriodo values (@Id, #@Fecha#, @EAProd, @MsProd, @MsCrecimiento, @MsDesarollo, @MsVaquillas, @MsSecas, @MsReto, @ValConsisitentes, @ValEAProd, @ValMsProd, @ValMsCrecimiento, @ValMsDesarollo, @ValMsVaquillas, @ValMsSecas, @ValMsReto)";

                query = query.Replace("@Id", valconAlim.Ran_ID.ToString())
                          .Replace("@Fecha", fecha.ToString("yyyy/MM/dd"))
                          .Replace("@EAProd", valconAlim.EAProd.ToString())
                          .Replace("@MsProd", valconAlim.MsProd.ToString())
                          .Replace("@MsCrecimiento", valconAlim.MsCrecimiento.ToString())
                          .Replace("@MsDesarollo", valconAlim.MsDesarollo.ToString())
                          .Replace("@MsVaquillas", valconAlim.MsVaquillas.ToString())
                          .Replace("@MsSecas", valconAlim.MsSecas.ToString())
                          .Replace("@MsReto", valconAlim.MsReto.ToString())
                          .Replace("@ValConsisitentes", valconAlim.ValConsisitentes.ToString())
                          .Replace("@ValEAProd", valconAlim.EAProdValor.ToString())
                          .Replace("@ValMsProd", valconAlim.MsProdValor.ToString())
                          .Replace("@ValMsCrecimiento", valconAlim.MsCrecimientosValor.ToString())
                          .Replace("@ValMsDesarollo", valconAlim.MsDesarolloValor.ToString())
                          .Replace("@ValMsVaquillas", valconAlim.MsVaquillasValor.ToString())
                          .Replace("@ValMsSecas", valconAlim.MsSecasValor.ToString())
                          .Replace("@ValMsReto", valconAlim.MsRetoValor.ToString());

                db.Conectar();
                db.CrearComando(query, tipoComandoAccess.query);
                int insert = db.EjecutarComando();
                Console.WriteLine("{0} Elemento Insertado", insert);

            }
            catch (Exception ex) { mensaje = ex.Message; }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }
        }

        private InventarioAfiXDia IventarioAFI(int ranId, DateTime fecha, ref string mensaje)
        {
            InventarioAfiXDia inventario = new InventarioAfiXDia();
            ModeloDatosSQL db = new ModeloDatosSQL(conexionSQL);
            mensaje = string.Empty;

            try
            {
                string query = @"SELECT ran_id 
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
                query = query.Replace("@FechaFinal", fecha.Date.ToString("yyyy/MM/dd"))
                             .Replace("@Ran_ID", ranId.ToString());

                db.Conectar();
                db.CrearComando(query, tipoComandoSQL.query);
                inventario = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new InventarioAfiXDia()
                {
                    Fecha = fecha,
                    Ordeño = x["Ordeño"] != DBNull.Value ? Convert.ToDecimal(x["Ordeño"]) : 0,
                    Secas = x["Secas"] != DBNull.Value ? Convert.ToDecimal(x["Secas"]) : 0,
                    Reto = x["Reto"] != DBNull.Value ? Convert.ToDecimal(x["Reto"]) : 0,
                    Jaulas = 0,
                    Crecimiento = x["DesteteUno"] != DBNull.Value ? Convert.ToDecimal(x["DesteteUno"]) : 0,
                    Desarrollo = x["DesteteDos"] != DBNull.Value ? Convert.ToDecimal(x["DesteteDos"]) : 0,
                    Vaquillas = x["Vaquillas"] != DBNull.Value ? Convert.ToDecimal(x["Vaquillas"]) : 0,
                    InventarioTotal = 0
                }).ToList().FirstOrDefault();

            }
            catch (Exception ex) { mensaje = ex.Message; }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return inventario;
        }

        private InventarioAfiXDia PromedioIventarioAFI(int ranId, DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            InventarioAfiXDia inventario = new InventarioAfiXDia();
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

            try
            {
                string query = @"SELECT  ROUND(AVG(d1))  AS Crecimiento
                               ,ROUND(AVG(d2))  AS Desarrollo
                               ,ROUND(AVG(v))  AS Vaquillas
                               ,ROUND(AVG(s)) AS Secas
                               ,ROUND(AVG(o)) AS Ordeño
                               ,ROUND(AVG(r)) AS Reto
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
                query = query.Replace("@fechaI", ConvertToJulian(fechaInicio).ToString()).Replace("@fechaF", ConvertToJulian(fechaFin).ToString());

                db.Conectar();
                db.CrearComando(query, tipoComandoAccess.query);
                inventario = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new InventarioAfiXDia()
                {
                    Ordeño = x["Ordeño"] != DBNull.Value ? Convert.ToDecimal(x["Ordeño"]) : 0,
                    Secas = x["Secas"] != DBNull.Value ? Convert.ToDecimal(x["Secas"]) : 0,
                    Reto = x["Reto"] != DBNull.Value ? Convert.ToDecimal(x["Reto"]) : 0,
                    Jaulas = 0,
                    Crecimiento = x["Crecimiento"] != DBNull.Value ? Convert.ToDecimal(x["Crecimiento"]) : 0,
                    Desarrollo = x["Desarrollo"] != DBNull.Value ? Convert.ToDecimal(x["Desarrollo"]) : 0,
                    Vaquillas = x["Vaquillas"] != DBNull.Value ? Convert.ToDecimal(x["Vaquillas"]) : 0,
                    InventarioTotal = 0
                }).ToList().FirstOrDefault();

            }
            catch (Exception ex) { mensaje = ex.Message; }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return inventario;
        }

        private PromedioMProducAlimentacion MProducAlim(DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            PromedioMProducAlimentacion response = new PromedioMProducAlimentacion();
            ModeloDatosAccess db = new ModeloDatosAccess(conexionAccess);
            mensaje = string.Empty;

            try
            {
                string query = @"SELECT  IIF(SUM(CONTEA) > 0,SUM(EA) / SUM(CONTEA),0)    AS EA2
                                        ,IIF(SUM(CONTMS) > 0,SUM(MS) / SUM(CONTMS),0)    AS MS2
                                        ,IIF(SUM(CONTDI) > 0,SUM(DI) / SUM(CONTDI),0)    AS DI2
                                        ,IIF(SUM(CONTDII) > 0,SUM(DII) / SUM(CONTDII),0) AS DII2
                                        ,IIF(SUM(CONTVQ) > 0,SUM(VQ) / SUM(CONTVQ),0)    AS VQ2
                                        ,IIF(SUM(CONTS) > 0,SUM(S) / SUM(CONTS),0)       AS S2
                                        ,IIF(SUM(CONTR) > 0,SUM(R) / SUM(CONTR),0)       AS R2
                                FROM
                                (
	                                SELECT  m.EAS                                                                       AS EA
	                                        ,IIF(m.EAS > 0,1,0)                                                          AS CONTEA
	                                        ,m.MSS                                                                       AS MS
	                                        ,IIF(MSS > 0,1,IIF(i.VACASORDEÑA > 0,1,0))                                   AS CONTMS
	                                        ,(m.MH_DI - m.SA_DI) * m.MS_DI                                               AS DI
	                                        ,IIF(((m.MH_DI - m.SA_DI) * m.MS_DI) > 0,1,IIF(i.destetadas > 0,1,0))        AS CONTDI
	                                        ,(m.MH_DII - m.SA_DII) * m.MS_DII                                            AS DII
	                                        ,IIF(((m.MH_DII - m.SA_DII) * m.MS_DII ) > 0,1,IIF(i.destetadas2 > 0,1,0))   AS CONTDII
	                                        ,((m.MH_VQ - m.SA_VQ) * m.MS_VQ)                                             AS VQ
	                                        ,IIF(((m.MH_VQ - m.SA_VQ) * m.MS_VQ) > 0,1,IIF(i.vaquillas > 0,1,0))         AS CONTVQ
	                                        ,((m.MH_S - m.SA_S) * m.MS_S)                                                AS S
	                                        ,IIF(((m.MH_S - m.SA_S) * m.MS_S) > 0,1,IIF(i.vacassecas > 0,1,0))           AS CONTS
	                                        ,((m.MH_R - m.SA_R) * m.MS_R)                                                AS R
	                                        ,IIF(((m.MH_R - m.SA_R) * m.MS_R) > 0,1,IIF((i.vqreto + i.vcreto ) > 0,1,0)) AS CONTR
	                                FROM MPRODUC m
	                                INNER JOIN INVENTARIOAFIXDIA i
	                                ON m.FECHA = i.FECHA
                                WHERE m.FECHA BETWEEN @fechaInicio AND @fechaFin ) Tabla";

                db.Conectar();
                db.CrearComando(query, tipoComandoAccess.query);
                db.AsignarParametro("@fechaInicio", ConvertToJulian(fechaInicio));
                db.AsignarParametro("@fechaFin", ConvertToJulian(fechaInicio));
                response = db.EjecutarConsultaTabla().AsEnumerable().Select(x => new PromedioMProducAlimentacion()
                {
                    EA = x["EA2"] != DBNull.Value ? Convert.ToDecimal(x["EA2"]) : 0,
                    Produccion = x["MS2"] != DBNull.Value ? Convert.ToDecimal(x["MS2"]) : 0,
                    Secas = x["S2"] != DBNull.Value ? Convert.ToDecimal(x["S2"]) : 0,
                    Reto = x["R2"] != DBNull.Value ? Convert.ToDecimal(x["R2"]) : 0,
                    Crecimiento = x["DI2"] != DBNull.Value ? Convert.ToDecimal(x["DI2"]) : 0,
                    Desarrollo = x["DII2"] != DBNull.Value ? Convert.ToDecimal(x["DII2"]) : 0,
                    Vaquillas = x["VQ2"] != DBNull.Value ? Convert.ToDecimal(x["VQ2"]) : 0,

                }).ToList().FirstOrDefault();

            }
            catch (Exception ex) { mensaje = ex.Message; }
            finally
            {
                if (db.isConnected)
                    db.Desconectar();
            }

            return response;
        }

        #endregion
    }



    internal class ConfiguracionRancho
    {
        public string Erp_id { get; set; }
        public int Emp_id { get; set; }
        public int Track_id { get; set; }
        public int Ran_sie { get; set; }
        public int Emp_prorrateo { get; set; }
    }

    internal class CentroDeCosto
    {
        public string CCO { get; set; }
    }

    internal class CCOGHT
    {
        public string CCO_Clave { get; set; }
        public string CCO_Nombre { get; set; }
    }
}
