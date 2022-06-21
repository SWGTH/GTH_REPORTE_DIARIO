using FirebirdSql.Data.FirebirdClient;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReportePeriodo
{
    public class Connection
    {
        FbConnection ConnFB = new FbConnection();
        SqlConnection ConnAlimento = new SqlConnection();
        SqlConnection connSIO = new SqlConnection();
        SqlConnection connSIE = new SqlConnection();
        OleDbConnection ConnIIE = new OleDbConnection();

        FbCommand cmdFB = new FbCommand();
        SqlCommand cmdSIO = new SqlCommand();
        SqlCommand cmdSIE = new SqlCommand();
        SqlCommand cmdAlimento = new SqlCommand();
        OleDbCommand cmdIIE = new OleDbCommand();

        public void Iniciar()
        {
            //DBAlimento
            string cadenaAlimento = ConfigurationManager.AppSettings["CadenaCosumo"];
            cadenaAlimento = cadenaAlimento.Replace("$usr", "sa").Replace("$pas", "CiaPrest_");
            ConnAlimento.ConnectionString = cadenaAlimento;
            cmdAlimento.Connection = ConnAlimento;
            //DBSIO
            string cadenaSIO = ConfigurationManager.AppSettings["CadenaSIO"];
            cadenaSIO = cadenaSIO.Replace("$usr", "sa").Replace("$pas", "CiaPrest_");
            connSIO.ConnectionString = cadenaSIO;
            cmdSIO.Connection = connSIO;
            //DBSIE
            string cadenaSIE = ConfigurationManager.AppSettings["CadenaSIE"];
            cadenaSIE = cadenaSIE.Replace("$usr", "sa").Replace("$pas", "CiaPrest_");
            connSIE.ConnectionString = cadenaSIE;
            cmdSIE.Connection = connSIE;

            string cadenaAcces = ConfigurationManager.AppSettings["connMOVSIO"];
            cadenaAcces = cadenaAcces.Replace("$pwd", "CiaPrest_");
            ConnIIE.ConnectionString = cadenaAcces;
            cmdIIE.Connection = ConnIIE;
            IniciarTracker();
        }

        public void IniciarTracker()
        {
            string rancho = GetRanchoCadena();
            string tracker = "Tracker" + rancho + ".FDB";
            string cadenaTracker = ConfigurationManager.AppSettings["CadenaTracker"];
            cadenaTracker = cadenaTracker.Replace("$user", "SYSDBA").Replace("$pwd", "masterkey").Replace("$tracker", tracker);            
            ConnFB.ConnectionString = cadenaTracker;
            cmdFB.Connection = ConnFB;
        }

        public string GetRanchoCadena()
        {
            string rancho = "";
            try
            {
                ConnIIE.Open();
                DataTable dt;
                string query = "SELECT RANCHOLOCAL FROM RANCHOLOCAL";
                OleDbDataAdapter da = new OleDbDataAdapter(query, ConnIIE);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];
                int ranId = Convert.ToInt32(dt.Rows[0][0]);
                rancho = ranId > 9 ? ranId.ToString() : "0" + ranId.ToString();
            }
            catch (DbException e) { MessageBox.Show(e.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            catch (Exception e) { MessageBox.Show(e.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally
            {
                ConnIIE.Close();
            }
            return rancho;
        }

        public int GetRancho()
        {
            int ran_id = 0;
            try
            {
                ConnIIE.Open();
                DataTable dt;
                string query = "SELECT RANCHOLOCAL FROM RANCHOLOCAL";
                OleDbDataAdapter da = new OleDbDataAdapter(query, ConnIIE);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];
                ran_id = Convert.ToInt32(dt.Rows[0][0]);                
            }
            catch (DbException e) { MessageBox.Show(e.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            catch (Exception e) { MessageBox.Show(e.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally
            {
                ConnIIE.Close();
            }
            return ran_id;
        }

        public void QueryAlimento(string query, out DataTable dt)
        {
            dt = new DataTable();

            try
            {
                ConnAlimento.Open();
                cmdAlimento.CommandText = "SET DATEFORMAT 'YMD'";
                cmdAlimento.ExecuteNonQuery();
                SqlDataAdapter da = new SqlDataAdapter(query, ConnAlimento);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];
            }
            catch (DbException e) { MessageBox.Show(e.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            catch (Exception e) { MessageBox.Show(e.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally
            {
                ConnAlimento.Close();
            }
        }

        public void QueryAlimento(string query)
        {
            try
            {
                ConnAlimento.Open();
                cmdAlimento.CommandText = "SET DATEFORMAT 'YMD'";
                cmdAlimento.ExecuteNonQuery();
                cmdAlimento.CommandText = query;
                cmdAlimento.ExecuteNonQuery();
                
            }
            catch (DbException e) { MessageBox.Show(e.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            catch (Exception e) { MessageBox.Show(e.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally
            {
                ConnAlimento.Close();
            }
        }

        public void QueryTracker(string query, out DataTable dt)
        {
            dt = new DataTable();
            try
            {
                ConnFB.Open();
                FbDataAdapter da = new FbDataAdapter(query, ConnFB);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];
            }
            catch (DbException e) { MessageBox.Show(e.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            catch (Exception e) { MessageBox.Show(e.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally
            {
                ConnFB.Close();
            }
        }

        public void QueryMovsio(string query, out DataTable dt)
        {
            dt = new DataTable();

            try
            {
                ConnIIE.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(query, ConnIIE);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];
            }
            catch (DbException e) { MessageBox.Show(e.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            catch (Exception e) { MessageBox.Show(e.ToString(), "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally
            {
                ConnIIE.Close();
            }
        }

        public int QueryMovsio(string query) 
        {
            int insert = 0;

            try 
            {
                ConnIIE.Open();
                cmdIIE.CommandText = query;
                insert = cmdIIE.ExecuteNonQuery();

            }
            catch(DbException ex) { Console.WriteLine(ex.Message); }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            finally
            {
                ConnIIE.Close();
            }

            return insert;
        }


        public void DeleteAlimento(string tabla, string condicion)
        {
            try
            {
                int deleteCount = 0;
                ConnAlimento.Open();
                cmdAlimento.CommandText = "SET DATEFORMAT 'YMD'";
                cmdAlimento.ExecuteNonQuery();
                string query = "delete from " + tabla + " " + condicion;
                cmdAlimento.CommandText = query;
                deleteCount = cmdAlimento.ExecuteNonQuery();
                Console.WriteLine("{0} registros eliminados en la tabla {1}", deleteCount, tabla);
            }
            catch (DbException e) { MessageBox.Show(e.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            catch (Exception e) { MessageBox.Show(e.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally { ConnAlimento.Close(); }
        }

        public void InsertMasivAlimento(string tabla, string valores)
        {
            try
            {
                int insertCount = 0;
                ConnAlimento.Open();
                cmdAlimento.CommandText = "SET DATEFORMAT 'YMD'";
                cmdAlimento.ExecuteNonQuery();
                string query = "INSERT INTO " + tabla + " VALUES " + valores;
                cmdAlimento.CommandTimeout = 120;
                cmdAlimento.CommandText = query;
                insertCount = cmdAlimento.ExecuteNonQuery();
                Console.WriteLine("{0} registros insertados en la tabla {1}", insertCount, tabla);
            }
            catch (DbException e) { MessageBox.Show(e.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            catch (Exception e) { MessageBox.Show(e.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally
            {
                ConnAlimento.Close();
            }
        }

        public void QuerySIO(string query, out DataTable dt)
        {
            dt = new DataTable();
            try
            {
                connSIO.Open();
                cmdSIO.CommandText = "SET DATEFORMAT 'YMD'";
                cmdSIO.ExecuteNonQuery();
                SqlDataAdapter da = new SqlDataAdapter(query, connSIO);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];
            }
            catch (DbException e) { MessageBox.Show(e.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            catch (Exception e) { MessageBox.Show(e.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally { connSIO.Close(); }
        }

        public int RanchoSie(int ran_id)
        {
            int ran_sie = 0;
            try
            {
                connSIO.Open();
                DataTable dt;
                string query = "SELECT ran_sie FROM configuracion WHERE ran_id = " + ran_id;
                SqlDataAdapter da = new SqlDataAdapter(query, connSIO);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];
                ran_sie = dt.Rows[0][0] != DBNull.Value ? Convert.ToInt32(dt.Rows[0][0]) : 0;
            }
            catch (DbException e) { MessageBox.Show(e.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            catch (Exception e) { MessageBox.Show(e.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally
            {
                connSIO.Close();
            }

            return ran_sie;
        }

        public void InsertSelecttAlimento(string query)
        {
            try
            {
                int insertCount = 0;
                ConnAlimento.Open();
                cmdAlimento.CommandText = "SET DATEFORMAT 'YMD'";
                cmdAlimento.ExecuteNonQuery();
                cmdAlimento.CommandText = query;
                insertCount = cmdAlimento.ExecuteNonQuery();
                Console.WriteLine("{0} registros insertados", insertCount);
            }
            catch (DbException ex) { MessageBox.Show(ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            catch (Exception ex) { MessageBox.Show(ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally { ConnAlimento.Close(); }
        }

        public void QueryMovGanado(string query, out DataTable dt)
        {
            dt = new DataTable();
            try
            {
                ConnIIE.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(query, ConnIIE);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];
            }
            catch (DbException e) { MessageBox.Show(e.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            catch (Exception e) { MessageBox.Show(e.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally
            {
                ConnIIE.Close();
            }
        }
    }
}
