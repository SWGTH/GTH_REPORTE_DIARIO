using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.Common;
using ReportePeriodo.Entidad;
using ReportePeriodo.Datos;
using LibreriaAlimentacion;
using System.ComponentModel;

namespace ReportePeriodo.Modelo
{
    public class Modelo1 : IModelo1
	{        
        public Modelo1()
        { 
        
        }

		private DatosTeorico DatosTeorico(int ranId, int horaCorte, string etapa, DateTime fecha, ref string mensaje)
		{
			DatosTeorico datosT = new DatosTeorico();
			ModeloDatos bd = new ModeloDatos();
			mensaje = string.Empty;
			string campoInventario = "";
			switch (etapa)
			{
				case "10,11,12,13":
					campoInventario = "ia_vacas_ord";
					break;
				case "21":
					campoInventario = "ia_vacas_secas";
					break;
				case "22":
					campoInventario = "ia_vqreto+ia_vcreto";
					break;
				case "31":
					campoInventario = "ia_jaulas";
					break;
				case "32":
					campoInventario = "ia_destetadas";
					break;
				case "33":
					campoInventario = "ia_destetadas2";
					break;
				case "34":
					campoInventario = "ia_vaquillas";
					break;

			}
			try
			{
				string query = @"SELECT  ROUND(SUM(MH),5)                                AS MH
									   ,ROUND(SUM(MS),5)                                AS MS
										,ROUND(IIF(SUM(MH) > 0,SUM(MS)/SUM(MH) * 100, 0),5) AS PorcMs
									   ,ROUND(SUM(COSTO),5)                             AS Costo
									   ,ROUND(IIF(SUM(MS) > 0,SUM(COSTO)/ SUM(MS),0),5) AS kgMS
								FROM
								(
									SELECT  ing_clave
										   ,ing_descripcion
										   ,IIF(vacas > 0,peso / vacas,0 )                           AS MH
										   ,(IIF(vacas > 0,peso / vacas,0 ) * ISNULL(PorcMs,0) /100) AS MS
										   ,IIF(vacas > 0,peso / vacas,0 ) * ISNULL(Costo.Costo,0)   AS COSTO
									FROM
									(
										SELECT  ran_id
											   ,ing_clave
											   ,ing_descripcion
											   ,SUM(peso_teorico) AS Peso
										FROM
										(
											SELECT  ran_id
												   ,ing_clave
												   ,ing_descripcion
												   ,peso_teorico
												   ,numvacas
											FROM racionTeorico
											WHERE ran_id = @rancho
											AND rac_fecha = @fecha
											AND CONVERT(int, SUBSTRING(racion_descripcion, 3, 2)) IN (@etapa)
											AND ing_clave <> ''
											UNION
											SELECT  ran_id
												   ,ing_clave
												   ,ing_descripcion
												   ,rac_mh
												   ,vacas
											FROM
											(
												SELECT  polvo.*
													   ,Case etp_id 
														WHEN 11 THEN (ia_lactancia1 - ia_vacassecasl1) 
														WHEN 12 THEN (ia_lactancia2 - ia_vacassecasl2) 
														WHEN 13 THEN ((ia_lactancia3 + ia_lactancia4) - (ia_vacassecasl4 + ia_vacassecasl4)) 
														WHEN 21 THEN ia_vacas_secas 
														WHEN 22 THEN (ia_vqreto + ia_vcreto) 
														WHEN 31 THEN ia_jaulas 
														WHEN 32 THEN ia_destetadas 
														WHEN 33 THEN ia_destetadas2 
														WHEN 34 THEN ia_vaquillas ELSE 0 End AS vacas
												FROM
												(
													SELECT  *
													FROM racion
													WHERE ran_id = @rancho
													AND ing_polvo = 1
													AND rac_fecha >= @fechaRacionIni
													AND rac_fecha < @fechaRacionFin 
												)polvo
												LEFT JOIN
												(
													SELECT  *
													FROM inventario_afi
													WHERE ia_fecha = @fecha
													AND ran_id = @rancho 
												) inventario
												ON polvo.ran_id = inventario.ran_id
											) polvo
											WHERE etp_id IN (@etapa)
											UNION
											SELECT  ran_id
												   ,ing_descripcion
												   ,ing_clave
												   ,SUM(peso_teorico) AS peso
												   ,SUM(numvacas)     AS vacas
											FROM
											(
												SELECT  ran_id
													   ,ing_clave
													   ,ing_descripcion
													   ,peso_teorico
													   ,numvacas
												FROM racionTeorico
												WHERE ran_id = @rancho
												AND rac_fecha = @fecha
												AND CONVERT(int, SUBSTRING(racion_descripcion, 3, 2)) IN (@etapa)
												AND ing_clave = '' 
											) otro
											GROUP BY  ran_id
													 ,ing_descripcion
													 ,ing_clave
										) Teorico
										GROUP BY  ran_id
												 ,ing_clave
												 ,ing_descripcion
									) Teorico
									LEFT JOIN
									(
										SELECT  ran_id
											   ,@campo AS Vacas
										FROM inventario_afi
										WHERE ia_fecha = @fecha
										AND ran_id = @rancho 
									) vacas
									ON Teorico.ran_id = vacas.ran_id
									LEFT JOIN
									(
										SELECT  Clave
											   ,Ingrediente
											   ,IIF(MH > 0,MS / MH * 100,0) AS PorcMs
										FROM
										(
											SELECT  R.Rancho
												   ,R.Clave
												   ,R.Ingrediente
												   ,SUM(R.PesoH) AS MH
												   ,SUM(R.PesoS) AS MS
											FROM
											(
												SELECT  ran_id            AS Rancho
													   ,r.ing_clave       AS Clave
													   ,r.ing_descripcion AS Ingrediente
													   ,SUM(rac_mh)       AS PesoH
													   ,SUM(rac_ms)       AS PesoS
												FROM racion r
												WHERE ran_id = @rancho
												AND rac_fecha >= @fechaRacionIni
												AND rac_fecha < @fechaRacionFin
												AND etp_id IN (@etapa)
												AND SUBSTRING(ing_clave, 1, 4) IN ('ALAS', 'ALFO')
												GROUP BY  ran_id
														 ,ing_clave
														 ,ing_descripcion
												UNION
												SELECT  T.Rancho
													   ,T.Clave
													   ,T.Ing
													   ,SUM(T.Peso)  AS MH
													   ,SUM(T.PesoS) AS MS
												FROM
												(
													SELECT  T1.Rancho
														   ,IIF(T2.Pmez IS NULL,T1.Clave,T2.Clave)              AS Clave
														   ,IIF(T2.Pmez IS NULL,T1.Ing,T2.Ing)                  AS Ing
														   ,IIF(T2.Pmez IS NULL,T1.Peso,T1.Peso * T2.Porc)      AS Peso
														   ,IIF(T2.Pmez IS NULL,T1.PesoS,T1.Peso * T2.PorcSeca) AS PesoS
													FROM
													(
														SELECT  T1.Rancho
															   ,T2.Clave
															   ,T2.Ing
															   ,(T1.Peso * T2.Porc)     AS Peso
															   ,(T1.Peso * T2.PorcSeca) AS PesoS
														FROM
														(
															SELECT  ran_id          AS Rancho
																   ,ing_descripcion AS Pmz
																   ,SUM(rac_mh)     AS Peso
																   ,SUM(rac_ms)     AS PesoS
															FROM racion
															WHERE ran_id = @rancho
															AND rac_fecha >= @fechaRacionIni
															AND rac_fecha < @fechaRacionFin
															AND etp_id IN (@etapa)
															AND ISNUMERIC(SUBSTRING(ing_descripcion, 1, 1)) > 0
															AND SUBSTRING(ing_descripcion, 3, 2) IN ('00', '01', '02')
															GROUP BY  ran_id
																	 ,ing_descripcion
														) T1
														LEFT JOIN
														(
															SELECT  pmez_descripcion     AS Pmez
																   ,ing_clave            AS Clave
																   ,ing_descripcion      AS Ing
																   ,pmez_porcentaje      AS Porc
																   ,pmez_porcentaje_seca AS PorcSeca
															FROM porcentaje_Premezcla
														)T2
														ON T1.Pmz = T2.Pmez
													) T1
													LEFT JOIN
													(
														SELECT  pmez_descripcion     AS Pmez
															   ,ing_clave            AS Clave
															   ,ing_descripcion      AS Ing
															   ,pmez_porcentaje      AS Porc
															   ,pmez_porcentaje_seca AS PorcSeca
														FROM porcentaje_Premezcla
													)T2
													ON T1.Ing = T2.Pmez
												) T
												GROUP BY  T.Rancho
														 ,T.Clave
														 ,T.Ing
											) R
											GROUP BY  R.Rancho
													 ,R.Clave
													 ,R.Ingrediente
										) PorcentajeMs
									)pms
									ON Teorico.ing_clave = pms.Clave AND Teorico.ing_descripcion = pms.Ingrediente
									LEFT JOIN
									(
										SELECT  Costos.ArticuloCve
											   ,IIF(SUM(Costos.Existencia) > 0 ,SUM((Costos.Existencia * Costos.Costo)) / SUM(Costos.Existencia),0) AS Costo
										FROM
										(
											SELECT  a.ran_id
												   ,ci.ArticuloCve
												   ,IIF(ci.Existencia = 0,1,ci.Existencia) AS Existencia
												   ,ci.Costo
												   ,ci.Fecha
											FROM DBALIMENTO.dbo.costo_ingrediente_h ci
											LEFT JOIN DBSIE.dbo.almacen a
											ON ci.AlmacenCve = a.alm_id
											WHERE CONVERT(INT, FORMAT(ci.Fecha, 'yyyyMM')) = @periodo 
										) Costos
										WHERE Costos.ran_id = @rancho
										GROUP BY  Costos.ArticuloCve
									)Costo
									ON Teorico.ing_clave = Costo.ArticuloCve
								) Teorico";

				query = query.Replace("@etapa", etapa).Replace("@campo", campoInventario);

				bd.Conectar();
				bd.CrearComando(query, tipoComando.query);
				bd.AsignarParametro("@rancho", ranId);
				bd.AsignarParametro("@periodo", Convert.ToInt32(fecha.ToString("yyyyMM")));
				bd.AsignarParametro("@fecha", fecha);
				bd.AsignarParametro("@fechaRacionIni", fecha.AddHours(horaCorte));
				bd.AsignarParametro("@fechaRacionFin", fecha.AddHours(24 + horaCorte));
				DataTable dt = bd.EjecutarConsultaTabla();

				if (dt.Rows.Count > 0)
				{
					datosT.FECHA = fecha;
					datosT.MH = dt.Rows[0]["MH"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["MH"]) : 0;
					datosT.MS = dt.Rows[0]["MS"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["MS"]) : 0;
					datosT.PORCENTAJE_MS = dt.Rows[0]["PorcMs"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["PorcMs"]) : 0;
					datosT.COSTO = dt.Rows[0]["Costo"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["Costo"]) : 0;
					datosT.KGMS = dt.Rows[0]["kgMS"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["kgMS"]) : 0;
				}
				else
				{
					datosT.FECHA = fecha;
				}
			}
			catch (DbException ex) { mensaje = ex.Message; }
			catch (Exception ex) { mensaje = ex.Message; }
			finally
			{
				if (bd.isConnected)
					bd.Desconectar();
			}

			return datosT;
		}
		/*
		private DatosTeorico DatosTeorico(int ranId, int horaCorte, string etapa, DateTime fecha, ref string mensaje)
        {
            DatosTeorico datosT = new DatosTeorico();
            ModeloDatos bd = new ModeloDatos();
            mensaje = string.Empty;

            try
            {
                string query = @"SELECT  SUM(Peso)                                                                        AS MH
										   ,SUM(Peso * PorcMs / 100)                                                         AS MS
										   ,SUM(Peso * Costo)                                                                AS Costo
										   ,IIF( SUM(Peso * PorcMs / 100) > 0,SUM(Peso * PorcMs / 100) / SUM(Peso) * 100,0)  AS PorcMs
										   ,IIF(SUM(Peso * PorcMs / 100) > 0,SUM(Peso * Costo) / SUM(Peso * PorcMs / 100),0) AS kgMS
									FROM
									(
										SELECT  Teorico.*
											   ,PorcentajeMs.PorcMs
											   ,Costo.Costo
										FROM
										(
											SELECT  ran_id
												   ,ing_clave AS Clave
												   ,SUM(peso) AS Peso
											FROM racionTeorico
											WHERE ran_id = @rancho
											AND rac_fecha = @fecha
											AND CONVERT(int, SUBSTRING(racion_descripcion, 3, 2)) IN (@etapa)
											AND ing_clave NOT LIKE ''
											GROUP BY  ran_id, ing_clave
											UNION
											SELECT  ran_id
												   ,ing_clave
												   ,SUM(peso)
											FROM
											(
												SELECT  ran_id
													   ,etp_id
													   ,ing_clave
													   ,IIF(vacas > 0,rac_mh / vacas ,0) AS peso
												FROM
												(
													SELECT  polvo.*
														   , Case etp_id 
					   											WHEN 11 THEN (ia_lactancia1 - ia_vacassecasl1) 
																WHEN 12 THEN (ia_lactancia2 - ia_vacassecasl2) 
																WHEN 13 THEN ((ia_lactancia3 + ia_lactancia4) - (ia_vacassecasl4 + ia_vacassecasl4)) 
																WHEN 21 THEN ia_vacas_secas 
																WHEN 22 THEN (ia_vqreto + ia_vcreto) 
																WHEN 31 THEN ia_jaulas WHEN 32 THEN ia_destetadas 
																WHEN 33 THEN ia_destetadas2 
																WHEN 34 THEN ia_vaquillas 
																ELSE 0 End AS vacas
													FROM
													(
														SELECT  *
														FROM racion
														WHERE ran_id = @rancho
														AND ing_polvo = 1
														AND rac_fecha >= @fechaRacionIni
														AND rac_fecha < @fechaRacionFin
													)polvo
													LEFT JOIN
													(
														SELECT  *
														FROM inventario_afi
														WHERE ia_fecha = @fecha
														AND ran_id = @rancho 
													) inventario ON polvo.ran_id = inventario.ran_id
												) polvo
												WHERE etp_id IN (@etapa) 
											)polvo
											GROUP BY  ran_id, ing_clave
										) Teorico
										LEFT JOIN
										(
											SELECT  Clave
												   ,IIF(MH > 0,MS / MH * 100,0) AS PorcMs
											FROM
											(
												SELECT  R.Rancho
													   ,R.Clave
													   ,R.Ingrediente
													   ,SUM(R.PesoH) AS MH
													   ,SUM(R.PesoS) AS MS
												FROM
												(
													SELECT  ran_id            AS Rancho
														   ,r.ing_clave       AS Clave
														   ,r.ing_descripcion AS Ingrediente
														   ,SUM(rac_mh)       AS PesoH
														   ,SUM(rac_ms)       AS PesoS
													FROM racion r
													WHERE ran_id IN (@rancho)
													AND rac_fecha >= @fechaRacionIni
													AND rac_fecha < @fechaRacionFin
													AND etp_id IN (@etapa)
													AND SUBSTRING(ing_clave, 1, 4) IN ('ALAS', 'ALFO')
													GROUP BY  ran_id, ing_clave, ing_descripcion
													UNION
													SELECT  T.Rancho
														   ,T.Clave
														   ,T.Ing
														   ,SUM(T.Peso)  AS MH
														   ,SUM(T.PesoS) AS MS
													FROM
													(
														SELECT  T1.Rancho
															   ,IIF(T2.Pmez IS NULL,T1.Clave,T2.Clave)              AS Clave
															   ,IIF(T2.Pmez IS NULL,T1.Ing,T2.Ing)                  AS Ing
															   ,IIF(T2.Pmez IS NULL,T1.Peso,T1.Peso * T2.Porc)      AS Peso
															   ,IIF(T2.Pmez IS NULL,T1.PesoS,T1.Peso * T2.PorcSeca) AS PesoS
														FROM
														(
															SELECT  T1.Rancho
																   ,T2.Clave
																   ,T2.Ing
																   ,(T1.Peso * T2.Porc)     AS Peso
																   ,(T1.Peso * T2.PorcSeca) AS PesoS
															FROM
															(
																SELECT  ran_id          AS Rancho
																	   ,ing_descripcion AS Pmz
																	   ,SUM(rac_mh)     AS Peso
																	   ,SUM(rac_ms)     AS PesoS
																FROM racion
																WHERE ran_id IN (@rancho)
																AND rac_fecha >= @fechaRacionIni
																AND rac_fecha < @fechaRacionFin
																AND etp_id IN (@etapa)
																AND ISNUMERIC(SUBSTRING(ing_descripcion, 1, 1)) > 0
																AND SUBSTRING(ing_descripcion, 3, 2) IN ('00', '01', '02')
																GROUP BY  ran_id, ing_descripcion
															) T1
															LEFT JOIN
															(
																SELECT  pmez_descripcion     AS Pmez
																	   ,ing_clave            AS Clave
																	   ,ing_descripcion      AS Ing
																	   ,pmez_porcentaje      AS Porc
																	   ,pmez_porcentaje_seca AS PorcSeca
																FROM porcentaje_Premezcla
															)T2 ON T1.Pmz = T2.Pmez
														) T1
														LEFT JOIN
														(
															SELECT  pmez_descripcion     AS Pmez
																   ,ing_clave            AS Clave
																   ,ing_descripcion      AS Ing
																   ,pmez_porcentaje      AS Porc
																   ,pmez_porcentaje_seca AS PorcSeca
															FROM porcentaje_Premezcla
														)T2 ON T1.Ing = T2.Pmez
													) T
													GROUP BY  T.Rancho, T.Clave, T.Ing
												) R
												GROUP BY  R.Rancho, R.Clave, R.Ingrediente
											) PorcentajeMs
										)PorcentajeMs ON Teorico.Clave = PorcentajeMs.Clave
										LEFT JOIN
										(
											SELECT  Costos.ArticuloCve
												   ,IIF(SUM(Costos.Existencia) > 0 ,SUM((Costos.Existencia * Costos.Costo)) / SUM(Costos.Existencia),0) AS Costo
											FROM
											(
												SELECT  a.ran_id
													   ,ci.ArticuloCve
													   ,IIF(ci.Existencia = 0,1,ci.Existencia) AS Existencia
													   ,ci.Costo
													   ,ci.Fecha
												FROM DBALIMENTO.dbo.costo_ingrediente_h ci
												LEFT JOIN DBSIE.dbo.almacen a
												ON ci.AlmacenCve = a.alm_id
												WHERE CONVERT(INT, FORMAT(ci.Fecha, 'yyyyMM')) = @periodo 
											) Costos
											WHERE Costos.ran_id = @rancho
											GROUP BY  Costos.ArticuloCve
										)Costo ON Teorico.Clave = Costo.ArticuloCve
									) T1";

				query = query.Replace("@etapa", etapa);

				bd.Conectar();
				bd.CrearComando(query, tipoComando.query);
				bd.AsignarParametro("@rancho", ranId);
				bd.AsignarParametro("@periodo", Convert.ToInt32(fecha.ToString("yyyyMM")));
				bd.AsignarParametro("@fecha", fecha);
				bd.AsignarParametro("@fechaRacionIni", fecha.AddHours(horaCorte));
				bd.AsignarParametro("@fechaRacionFin", fecha.AddHours(24 + horaCorte));
				DataTable dt = bd.EjecutarConsultaTabla();

				if (dt.Rows.Count > 0)
				{
					datosT.FECHA = fecha;
					datosT.MH = dt.Rows[0]["MH"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["MH"]) : 0;
					datosT.MS = dt.Rows[0]["MS"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["MS"]) : 0;
					datosT.PORCENTAJE_MS = dt.Rows[0]["PorcMs"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["PorcMs"]) : 0;
					datosT.COSTO = dt.Rows[0]["Costo"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["Costo"]) : 0;
					datosT.KGMS = dt.Rows[0]["kgMS"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["kgMS"]) : 0;					
				}
                else
                {
					datosT.FECHA = fecha;
				}
            }
            catch (DbException ex) { mensaje = ex.Message; }
            catch(Exception ex) { mensaje = ex.Message; }
            finally 
            {
                if (bd.isConnected)
                    bd.Desconectar();
            }

            return datosT;
        }
		*/

		public void DatosTeoricos(int ranId, int horaCorte, DateTime fechaIni, DateTime fechaFin, out List<DatosTeorico> listaProduccion, 
			out List<DatosTeorico> listaSecas, out List<DatosTeorico> listaReto, out List<DatosTeorico> listaDestet1, 
			out List<DatosTeorico> listaDestete2, out List<DatosTeorico> listaVaquillas, ref string mensaje)
		{
			listaProduccion = new List<DatosTeorico>();
			listaSecas = new List<DatosTeorico>();
			listaReto = new List<DatosTeorico>();
			listaDestet1 = new List<DatosTeorico>();
			listaDestete2 = new List<DatosTeorico>();
			listaVaquillas = new List<DatosTeorico>();
			mensaje = string.Empty;
			List<DateTime> listaFechas = ListaFechas(fechaIni, fechaFin);

			foreach (DateTime fecha in listaFechas)
			{
				GTH.GenerarCargaPremezcla(ranId.ToString(), "10,11,12,13,21,22,31,32,33,34", horaCorte, fecha, fecha, ref mensaje);
				
				DatosTeorico produccion = DatosTeorico(ranId, horaCorte, "10,11,12,13", fecha ,ref mensaje);
				DatosTeorico secas = DatosTeorico(ranId, horaCorte, "21", fecha, ref mensaje);
				DatosTeorico reto = DatosTeorico(ranId, horaCorte, "22", fecha, ref mensaje);
				DatosTeorico destete1 = DatosTeorico(ranId, horaCorte, "32", fecha, ref mensaje);
				DatosTeorico destete2 = DatosTeorico(ranId, horaCorte, "33", fecha, ref mensaje);
				DatosTeorico vaquillas = DatosTeorico(ranId, horaCorte, "34", fecha, ref mensaje);

				listaProduccion.Add(produccion);
				listaSecas.Add(secas);
				listaReto.Add(reto);
				listaDestet1.Add(destete1);
				listaDestete2.Add(destete2);
				listaVaquillas.Add(vaquillas);
			}

			DataTable dtProd = ListToDataTable(listaProduccion);
			DataTable dtSecas = ListToDataTable(listaSecas);
			DataTable dtReto = ListToDataTable(listaReto);
			DataTable dtD1 = ListToDataTable(listaDestet1);
			DataTable dtD2 = ListToDataTable(listaDestete2);
			DataTable dtVP = ListToDataTable(listaVaquillas);

		}

		public List<DateTime> ListaFechas(DateTime inicio, DateTime fin)
		{
			List<DateTime> listaFechas = new List<DateTime>();

			DateTime temp = inicio;
			while (temp <= fin)
			{
				listaFechas.Add(temp);
				temp = temp.AddDays(1);
			}

			return listaFechas;
		}

		public DatosProduccion DatosProduccion(DateTime inicio, DateTime fin, ref string mensaje)
		{
			DatosProduccion dp = new DatosProduccion();
			ModeloDatosAccess bd = new ModeloDatosAccess();
			mensaje = string.Empty;
			int julianaInicio = ConvertToJulian(inicio);
			int julianaFin = ConvertToJulian(fin);

			try
			{
				string query = @"SELECT  IIF(SUM(vacasordeña) > 0,ROUND(SUM(lecprod + antprod)/SUM(vacasordeña),4),0) AS MEDIA
									   ,AVG(LECPROD)                                                                 AS LECHE
									   ,AVG((LECPROD + ANTPROD))                                                     AS TOTAL
								FROM mproduc m
								LEFT JOIN inventario i ON m.fecha = i.fecha
								WHERE m.fecha BETWEEN @inicio AND @fin";

				bd.Conectar();
				bd.CrearComando(query, tipoComandoDB.query);
				bd.AsignarParametro("@inicio", julianaInicio);
				bd.AsignarParametro(" @fin", julianaFin);
				DataTable dt = bd.EjecutarConsultaTabla();

				if (dt.Rows.Count > 0)
				{
					dp.MEDIA = dt.Rows[0]["MEDIA"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["MEDIA"]) : 0;
					dp.LECHE = dt.Rows[0]["LECHE"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["LECHE"]) : 0;
					dp.TOTAL = dt.Rows[0]["TOTAL"] != DBNull.Value ? Convert.ToDecimal(dt.Rows[0]["TOTAL"]) : 0;
				}
			}
			catch (DbException ex) { mensaje = ex.Message; }
			catch (Exception ex) { mensaje = ex.Message; }
			finally
			{
				if (bd.isConnected)
					bd.Desconectar();
			}

			return dp;
		}

		public List<LecheBacteriologia> ListaLecheBacteriologia(DateTime inicio, DateTime fin, ref string mensaje)
		{
			List<LecheBacteriologia> lista = new List<LecheBacteriologia>();
			ModeloDatosAccess bd = new ModeloDatosAccess();
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
				query = query.Replace("@fechaIni", inicio.ToString("yyyy/MM/dd")).Replace("@fechaFin", fin.ToString("yyyy/MM/dd"));


				bd.Conectar();
				bd.CrearComando(query, tipoComandoDB.query);				
				lista = bd.EjecutarConsultaTabla().AsEnumerable().Select(x => new LecheBacteriologia()
				{ 
					FECHA = x["FECHA"] != DBNull.Value ? Convert.ToDateTime(x["FECHA"]) : new DateTime(1,1,1),
					BACTEROLOGIA = x["PRECIO_BACTERIOLOGIA"] != DBNull.Value ? Convert.ToDecimal(x["PRECIO_BACTERIOLOGIA"]) : 0
				}).ToList();
			}
			catch (DbException ex) { mensaje = ex.Message; }
			catch (Exception ex) { mensaje = ex.Message; }
			finally
			{
				if (bd.isConnected)
					bd.Desconectar();
			}

			return lista;
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

		private static int ConvertToJulian(DateTime Date)
		{
			TimeSpan ts = (Date - Convert.ToDateTime("01/01/1900"));
			int julianday = ts.Days + 2;
			return julianday;
		}
	}
}
