using System;
using System.Collections.Generic;
using System.Linq;
using ReportePeriodo.Entidad;
using ReportePeriodo.Datos;
using System.Data;
using gth = LibreriaAlimentacion.Entidad;
using LibreriaAlimentacion;


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
        gth.IndicadorReportePeriodo promedioCrecimiento;
        gth.IndicadorReportePeriodo promedioDesarrollo;
        gth.IndicadorReportePeriodo promedioVaquillas;

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
            promedioCrecimiento = new gth.IndicadorReportePeriodo();
            promedioDesarrollo = new gth.IndicadorReportePeriodo();
            promedioVaquillas = new gth.IndicadorReportePeriodo();
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

        public List<Hoja1> ReporteHoja1(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, out List<CalostroYOrdeña> datosCalostro, ref string mensaje)
        {
            List<Hoja1> response = new List<Hoja1>();
            mensaje = string.Empty;
            datosCalostro = new List<CalostroYOrdeña>();

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
                    hoja.Porcentaje_MS = hoja.MH != null && hoja.MS != null ? hoja.MH != 0 ? hoja.MS / hoja.MH * 100 : null : null;
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
            GTH.IndicadoresTeoricosxEstablo(rancho.Ran_ID, "21", fechaFin, out indicadorSecas, ref mensaje);
            GTH.IndicadoresTeoricosxEstablo(rancho.Ran_ID, "22", fechaFin, out indicadorReto, ref mensaje);
            GTH.IndicadoresTeoricosxEstablo(rancho.Ran_ID, "32", fechaFin, out indicadorCrecimiento, ref mensaje);
            GTH.IndicadoresTeoricosxEstablo(rancho.Ran_ID, "33", fechaFin, out indicadorDesarrollo, ref mensaje);
            GTH.IndicadoresTeoricosxEstablo(rancho.Ran_ID, "34", fechaFin, out indicadorVaquillas, ref mensaje);
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
                GTH.ReportePeriodo(rancho.Ran_ID.ToString(), rancho.TimeShiftTracker, "32", fechaInicio, fechaFin, out ingredientesAux, out promedioCrecimiento, out sobranteAux);
                GTH.ReportePeriodo(rancho.Ran_ID.ToString(), rancho.TimeShiftTracker, "33", fechaInicio, fechaFin, out ingredientesAux, out promedioDesarrollo, out sobranteAux);
                GTH.ReportePeriodo(rancho.Ran_ID.ToString(), rancho.TimeShiftTracker, "34", fechaInicio, fechaFin, out ingredientesAux, out promedioVaquillas, out sobranteAux);
            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
        }

        public Hoja1 PromedioHoja1(Rancho rancho, DateTime fechaInicio, DateTime fechaFin, ref string mensaje)
        {
            Hoja1 response = new Hoja1() { Dia = "PROM" };
            PromedioHoja1 promH1 = rancho.No_ID_Real ? PromedioHoja1NoId(fechaInicio, fechaFin, ref mensaje) : PromedioHoja1NoId(fechaInicio, fechaFin, ref mensaje);
            mensaje = string.Empty;

            try
            {
                CalostroYOrdeña promedioCalostro = PromedioCalostro(fechaInicio, fechaFin);

                response.Ordeño = promH1.Ordeño;
                response.Secas = promH1.Secas;
                response.Hato = promH1.Hato;
                response.Porcentaje_Lact = promH1.Porcentaje_Lact;
                response.Porcentaje_Prot = promH1.Porcentaje_Prot;
                response.Urea = promH1.Urea;
                response.Porcentaje_Grasa = promH1.Porcentaje_Grasa;
                response.CCS = promH1.CCS;
                response.CTD = promH1.CTD;
                response.DEC = null;

                response.Leche = promH1.Leche;
                response.Antib = promH1.Antib;
                response.Media = (decimal)promedioProduccion.MEDIA;
                response.Total = promH1.Total;
                response.DEL = promH1.DEL;
                response.ANT = promH1.Ant;

                response.EA = (decimal)promedioProduccion.EA;
                response.ILCA = (decimal)promedioProduccion.ILCA_PRODUCCION;
                response.IC = (decimal)promedioProduccion.IC_PRODUCCION;
                response.Costo_Litro = (decimal)promedioProduccion.PRECIOL;
                response.MH = (decimal)promedioProduccion.MH;
                response.Porcentaje_MS = (decimal)promedioProduccion.PORCENTAJEMS;
                response.MS = (decimal)promedioProduccion.MS;
                response.SA = (decimal)promedioProduccion.SA;
                response.MSS = (decimal)promedioProduccion.MSS;
                response.EAS = (decimal)promedioProduccion.EAS;
                response.Porcentaje_Sob = (decimal)promedioProduccion.EA;
                response.Costo_Prod = (decimal)promedioProduccion.COSTO;
                response.Costo_MS = (decimal)promedioProduccion.PRECIOKGMS;

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

                foreach (Hoja1 item in reporte)
                {
                    busquedaCalostro = (from x in datosCalostro where x.Fecha.Day.ToString() == item.Dia select x).ToList().FirstOrDefault();
                    busquedaBacteriologia = (from x in datosBacteriologia where x.FECHA.Day.ToString() == item.Dia select x).ToList().FirstOrDefault();

                    item.Color_ILCA = Color1Hoja1(item.ILCA, Convert.ToDecimal(promedioProduccion.ILCA_PRODUCCION));
                    item.Color_IC = Color1Hoja1(item.ILCA, Convert.ToDecimal(promedioProduccion.IC_PRODUCCION));
                    item.Color_CostoLitro = Color2Hoja1(item.Costo_Litro, promedio.Costo_Litro);
                    item.Color_MH = ColorDato(item.Ordeño, item.MH, promedio.MH);
                    item.Color_PorcentajeMs = ColorDato(item.Ordeño, item.MH, promedio.Porcentaje_MS);
                    item.Color_MS = ColorDato(item.Ordeño, item.MH, promedio.MS);
                    item.Color_CostoProd = ColorDato(item.Ordeño, item.MH, promedio.Costo_Prod);
                    item.Color_CostoMs = ColorDato(item.Ordeño, item.MH, promedio.Costo_MS);
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
                porcentaje.Ordeño = diferencia.Ordeño != null && añoAnterior.Ordeño != null && añoAnterior.Ordeño != 0 ? diferencia.Ordeño / añoAnterior.Ordeño * 100 : 0;
                porcentaje.Secas = diferencia.Secas != null && añoAnterior.Secas != null && añoAnterior.Secas != 0 ? diferencia.Secas / añoAnterior.Secas * 100 : 0;
                porcentaje.Hato = diferencia.Hato != null && añoAnterior.Hato != null && añoAnterior.Hato != 0 ? diferencia.Hato / añoAnterior.Hato * 100 : 0;
                porcentaje.Porcentaje_Lact = diferencia.Porcentaje_Lact != null && añoAnterior.Porcentaje_Lact != null && añoAnterior.Porcentaje_Lact != 0 ? diferencia.Porcentaje_Lact / añoAnterior.Porcentaje_Lact * 100 : 0;
                porcentaje.Porcentaje_Prot = diferencia.Porcentaje_Prot != null && añoAnterior.Porcentaje_Prot != null && añoAnterior.Porcentaje_Prot != 0 ? diferencia.Porcentaje_Prot / añoAnterior.Porcentaje_Prot * 100 : 0;
                porcentaje.Urea = diferencia.Urea != null && añoAnterior.Urea != null && añoAnterior.Urea != 0 ? diferencia.Urea / añoAnterior.Urea * 100 : 0;
                porcentaje.Porcentaje_Grasa = diferencia.Porcentaje_Grasa != null && añoAnterior.Porcentaje_Grasa != null && añoAnterior.Porcentaje_Grasa != 0 ? diferencia.Porcentaje_Grasa / añoAnterior.Porcentaje_Grasa * 100 : 0;
                porcentaje.CCS = diferencia.CCS != null && añoAnterior.CCS != null && añoAnterior.CCS != 0 ? diferencia.CCS / añoAnterior.CCS * 100 : 0;
                porcentaje.CTD = diferencia.CTD != null && añoAnterior.CTD != null && añoAnterior.CTD != 0 ? diferencia.CTD / añoAnterior.CTD * 100 : 0;
                porcentaje.DEC = null;

                porcentaje.Leche = diferencia.Leche != null && añoAnterior.Leche != null && añoAnterior.Leche != 0 ? diferencia.Leche / añoAnterior.Leche * 100 : 0;
                porcentaje.Antib = diferencia.Antib != null && añoAnterior.Antib != null && añoAnterior.Antib != 0 ? diferencia.Antib / añoAnterior.Antib * 100 : 0;
                porcentaje.Media = diferencia.Media != null && añoAnterior.Media != null && añoAnterior.Media != 0 ? diferencia.Media / añoAnterior.Media * 100 : 0;
                porcentaje.Total = diferencia.Total != null && añoAnterior.Total != null && añoAnterior.Total != 0 ? diferencia.Total / añoAnterior.Total * 100 : 0;
                porcentaje.DEL = diferencia.DEL != null && añoAnterior.DEL != null && añoAnterior.DEL != 0 ? diferencia.DEL / añoAnterior.DEL * 100 : 0;
                porcentaje.ANT = diferencia.ANT != null && añoAnterior.ANT != null && añoAnterior.ANT != 0 ? diferencia.ANT / añoAnterior.ANT * 100 : 0;

                porcentaje.EA = diferencia.EA != null && añoAnterior.EA != null && añoAnterior.EA != 0 ? diferencia.EA / añoAnterior.EA * 100 : 0;
                porcentaje.ILCA = diferencia.ILCA != null && añoAnterior.ILCA != null && añoAnterior.ILCA != 0 ? diferencia.ILCA / añoAnterior.ILCA * 100 : 0;
                porcentaje.IC = diferencia.IC != null && añoAnterior.IC != null && añoAnterior.IC != 0 ? diferencia.IC / añoAnterior.IC * 100 : 0;
                porcentaje.Costo_Litro = diferencia.Costo_Litro != null && añoAnterior.Costo_Litro != null && añoAnterior.Costo_Litro != 0 ? diferencia.Costo_Litro / añoAnterior.Costo_Litro * 100 : 0;
                porcentaje.MH = diferencia.MH != null && añoAnterior.MH != null && añoAnterior.MH != 0 ? diferencia.MH / añoAnterior.MH * 100 : 0;
                porcentaje.Porcentaje_MS = diferencia.Porcentaje_MS != null && añoAnterior.Porcentaje_MS != null && añoAnterior.Porcentaje_MS != 0 ? diferencia.Porcentaje_MS / añoAnterior.Porcentaje_MS * 100 : 0;
                porcentaje.MS = diferencia.MS != null && añoAnterior.MS != null && añoAnterior.MS != 0 ? diferencia.MS / añoAnterior.MS * 100 : 0;
                porcentaje.SA = diferencia.SA != null && añoAnterior.SA != null && añoAnterior.SA != 0 ? diferencia.SA / añoAnterior.SA * 100 : 0;
                porcentaje.MSS = diferencia.MSS != null && añoAnterior.MSS != null && añoAnterior.MSS != 0 ? diferencia.MSS / añoAnterior.MSS * 100 : 0;
                porcentaje.EAS = diferencia.EAS != null && añoAnterior.EAS != null && añoAnterior.EAS != 0 ? diferencia.EAS / añoAnterior.EAS * 100 : 0;
                porcentaje.Porcentaje_Sob = diferencia.Porcentaje_Sob != null && añoAnterior.Porcentaje_Sob != null && añoAnterior.Porcentaje_Sob != 0 ? diferencia.Porcentaje_Sob / añoAnterior.Porcentaje_Sob * 100 : 0;
                porcentaje.Costo_Prod = diferencia.Costo_Prod != null && añoAnterior.Costo_Prod != null && añoAnterior.Costo_Prod != 0 ? diferencia.Costo_Prod / añoAnterior.Costo_Prod * 100 : 0;
                porcentaje.Costo_MS = diferencia.Costo_MS != null && añoAnterior.Costo_MS != null && añoAnterior.Costo_MS != 0 ? diferencia.Costo_MS / añoAnterior.Costo_MS * 100 : 0;

                porcentaje.Cribas_N1 = diferencia.Cribas_N1 != null && añoAnterior.Cribas_N1 != null && añoAnterior.Cribas_N1 != 0 ? diferencia.Cribas_N1 / añoAnterior.Cribas_N1 * 100 : 0;
                porcentaje.Cribas_N2 = diferencia.Cribas_N2 != null && añoAnterior.Cribas_N2 != null && añoAnterior.Cribas_N2 != 0 ? diferencia.Cribas_N2 / añoAnterior.Cribas_N2 * 100 : 0;
                porcentaje.Cribas_N3 = diferencia.Cribas_N3 != null && añoAnterior.Cribas_N3 != null && añoAnterior.Cribas_N3 != 0 ? diferencia.Cribas_N3 / añoAnterior.Cribas_N3 * 100 : 0;
                porcentaje.Cribas_N4 = diferencia.Cribas_N4 != null && añoAnterior.Cribas_N4 != null && añoAnterior.Cribas_N4 != 0 ? diferencia.Cribas_N4 / añoAnterior.Cribas_N4 * 100 : 0;

                porcentaje.NoID_Ses1 = diferencia.NoID_Ses1 != null && añoAnterior.NoID_Ses1 != null && añoAnterior.NoID_Ses1 != 0 ? diferencia.NoID_Ses1 / añoAnterior.NoID_Ses1 * 100 : 0;
                porcentaje.NoID_Ses2 = diferencia.NoID_Ses2 != null && añoAnterior.NoID_Ses2 != null && añoAnterior.NoID_Ses2 != 0 ? diferencia.NoID_Ses2 / añoAnterior.NoID_Ses2 * 100 : 0;
                porcentaje.NoID_Ses3 = diferencia.NoID_Ses3 != null && añoAnterior.NoID_Ses3 != null && añoAnterior.NoID_Ses3 != 0 ? diferencia.NoID_Ses3 / añoAnterior.NoID_Ses3 * 100 : 0;

                porcentaje.SES1 = diferencia.SES1 != null && añoAnterior.SES1 != null && añoAnterior.SES1 != 0 ? diferencia.SES1 / añoAnterior.SES1 * 100 : 0;
                porcentaje.SES2 = diferencia.SES2 != null && añoAnterior.SES2 != null && añoAnterior.SES2 != 0 ? diferencia.SES2 / añoAnterior.SES2 * 100 : 0;
                porcentaje.SES3 = diferencia.SES3 != null && añoAnterior.SES3 != null && añoAnterior.SES3 != 0 ? diferencia.SES3 / añoAnterior.SES3 * 100 : 0;

                diferencia.Porcentaje_Revueltas = diferencia.Porcentaje_Revueltas != null && añoAnterior.Porcentaje_Revueltas != null && añoAnterior.Porcentaje_Revueltas != 0 ? diferencia.Porcentaje_Revueltas / añoAnterior.Porcentaje_Revueltas * 100 : 0;
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
                    Litros_Sesion3 = x["LITROS3"] != DBNull.Value ? Convert.ToDecimal(x["LITROS3"]) : 0
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
                string query = @"SELECT  IIF(ISNULL(AVG(T3.TOTAL)),NULL,CLng(AVG(T3.TOTAL)))
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
                string query = @"SELECT  IIF(ISNULL(AVG(T3.TOTAL)),NULL,CLng(AVG(T3.TOTAL)))
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
                if (valor != null)
                {
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


        #endregion

    }
}
