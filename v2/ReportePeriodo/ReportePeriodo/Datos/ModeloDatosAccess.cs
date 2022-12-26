using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportePeriodo.Datos
{
    public enum tipoComandoDB
    {
        query,
        storedProcedure
    }

    public class ModeloDatosAccess
    {
        private OleDbConnection Conexion = null;
        private OleDbCommand Comando = null;
        private OleDbDataAdapter Adapter = null;
        private OleDbTransaction Transaccion = null;
        private DataTable Tabla = null;
        private string cadenaConexion = string.Empty;

        public ModeloDatosAccess()
        {
            Configurar();
        }

        //Propiedades
        public bool isInicializate
        {
            get
            {
                if (this.Conexion == null)
                    return false;
                return true;
            }
        }

        public bool isConnected
        {
            get
            {
                if (this.Conexion != null && !this.Conexion.State.Equals(ConnectionState.Closed))
                    return true;
                return false;
            }
        }

        /// <summary>
        /// Configura el acceso a la base de datos para su utilización.
        /// </summary>
        /// <exception cref="BaseDatosException">Si existe un error al cargar la configuración.</exception>
        private void Configurar()
        {
            try
            {
                this.cadenaConexion = @"Provider = Microsoft.Jet.OLEDB.4.0; Data Source = C:\MOVGANADO\movsio.mdb; JET OLEDB:DATABASE PASSWORD = CiaPrest_";
            }
            catch (ConfigurationException ex)
            {
                throw new Exception("Error al cargar la configuración del acceso a datos.", ex);
            }
        }

        /// <summary>
        /// Permite desconectarse de la base de datos.
        /// </summary>
        public void Desconectar()
        {
            if (this.Conexion.State.Equals(ConnectionState.Open))
            {
                this.Conexion.Close();
            }
        }

        /// <summary>
        /// Se concecta con la base de datos.
        /// </summary>
        /// <exception cref="BaseDatosException">Si existe un error al conectarse.</exception>
        public void Conectar()
        {
            if (this.Conexion != null && !this.Conexion.State.Equals(ConnectionState.Closed))
            {
                throw new Exception("La conexión ya se encuentra abierta.");
            }
            try
            {
                if (this.Conexion == null)
                {
                    this.Conexion = new OleDbConnection();
                    this.Conexion.ConnectionString = cadenaConexion;
                }
                this.Conexion.Open();
            }
            catch (DataException ex)
            {
                throw new Exception("Error al conectarse a la base de datos.", ex);
            }
        }

        /// <summary>
        /// Crea un comando en base a una sentencia SQL.
        /// Ejemplo:
        /// <code>SELECT * FROM Tabla WHERE campo1=@campo1, campo2=@campo2</code>
        /// Guarda el comando para el seteo de parámetros y la posterior ejecución.
        /// </summary>
        /// <param name="sentenciaSQL">La sentencia SQL con el formato: SENTENCIA [param = @param,]</param>
        public void CrearComando(string sentenciaSQL, tipoComandoDB tipo)
        {
            this.Comando = new OleDbCommand();
            this.Comando.Connection = Conexion;
            this.Comando.CommandTimeout = 0;
            if (tipo == tipoComandoDB.query)
            {
                this.Comando.CommandType = CommandType.Text;
            }
            else
            {
                this.Comando.CommandType = CommandType.StoredProcedure;
            }
            this.Comando.CommandText = sentenciaSQL;
            if (this.Transaccion != null)
            {
                this.Comando.Transaction = this.Transaccion;
            }

        }

        /// <summary>
        /// Asigna un parámetro de tipo bytes al comando creado.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        /// <param name="valor">El valor del parámetro.</param>
        public void AsignarParametro(string nombre, byte[] valor)
        {
            Comando.Parameters.AddWithValue(nombre, valor);
        }

        public void AsignarParametro(string nombre, int valor, ParameterDirection direction)
        {
            OleDbParameter sqlPara = new OleDbParameter(nombre, OleDbType.Integer, 2);
            sqlPara.Value = valor;
            sqlPara.Direction = direction;
            Comando.Parameters.Add(sqlPara);
        }

        public void AsignarParametro(string nombre, string valor, ParameterDirection direction)
        {
            OleDbParameter sqlPara = new OleDbParameter(nombre, OleDbType.VarChar, 20);
            sqlPara.Value = valor;
            sqlPara.Direction = direction;
            Comando.Parameters.Add(sqlPara);
        }

        /// <summary>
        /// Asigna un parámetro de tipo cadena al comando creado.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        /// <param name="valor">El valor del parámetro.</param>
        public void AsignarParametro(string nombre, string valor)
        {
            Comando.Parameters.AddWithValue(nombre, valor);
        }

        /// <summary>
        /// Asigna un parámetro de tipo entero al comando creado.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        /// <param name="valor">El valor del parámetro.</param>
        public void AsignarParametro(string nombre, int valor)
        {
            Comando.Parameters.AddWithValue(nombre, valor);
        }

        /// <summary>
        /// Asigna un parámetro de tipo entero al comando creado.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        /// <param name="valor">El valor del parámetro.</param>
        public void AsignarParametro(string nombre, long valor)
        {
            Comando.Parameters.AddWithValue(nombre, valor);
        }

        /// <summary>
        /// Asigna un parámetro de double entero al comando creado.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        /// <param name="valor">El valor del parámetro.</param>
        public void AsignarParametro(string nombre, double valor)
        {
            Comando.Parameters.AddWithValue(nombre, valor);
        }

        public void AsignarParametro(string nombre, System.Data.SqlTypes.SqlInt32 valor)
        {
            Comando.Parameters.AddWithValue(nombre, valor);
        }


        public void AsignarParametro(string nombre, System.Data.SqlTypes.SqlDouble valor)
        {
            Comando.Parameters.AddWithValue(nombre, valor);
        }

        public void AsignarParametro(string nombre, System.Data.SqlTypes.SqlDecimal valor)
        {
            Comando.Parameters.AddWithValue(nombre, valor);
        }

        public void AsignarParametro(string nombre, System.Data.SqlTypes.SqlString valor)
        {
            Comando.Parameters.AddWithValue(nombre, valor);
        }

        /// <summary>
        /// Asigna un parámetro de tipo fecha al comando creado.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        /// <param name="valor">El valor del parámetro.</param>
        public void AsignarParametroFecha(string nombre, DateTime valor)
        {
            Comando.Parameters.AddWithValue(nombre, valor);
        }


        /// <summary>
        /// Asigna un parámetro de tipo fecha al comando creado.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        /// <param name="valor">El valor del parámetro.</param>
        public void AsignarParametroFecha(string nombre, System.Data.SqlTypes.SqlDateTime a)
        {
            Comando.Parameters.AddWithValue(nombre, a);
        }

        /// <summary>
        /// Asigna un parámetro de tipo fecha al comando creado.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        /// <param name="valor">El valor del parámetro.</param>
        public void AsignarParametro(string nombre, object valor)
        {
            Comando.Parameters.AddWithValue(nombre, valor);
        }

        /// <summary>
        /// Ejecuta el comando creado y retorna el resultado de la consulta.
        /// </summary>
        /// <returns>El resultado de la consulta.</returns>
        /// <exception cref="BaseDatosException">Si ocurre un error al ejecutar el comando.</exception>
        public OleDbDataReader EjecutarConsulta()
        {
            return this.Comando.ExecuteReader();
        }

        /// <summary>
        /// Ejecuta el comando creado y retorna un escalar.
        /// </summary>
        /// <returns>El escalar que es el resultado del comando.</returns>
        /// <exception cref="BaseDatosException">Si ocurre un error al ejecutar el comando.</exception>
        public int EjecutarEscalar()
        {
            int escalar = 0;
            object objeto = null;
            try
            {
                objeto = this.Comando.ExecuteScalar();
                if (objeto != null)
                {
                    if (!(objeto is bool))
                        return Convert.ToInt32(objeto);
                    else
                        int.TryParse(objeto.ToString(), out escalar);
                }
                // if (objecto != DBNull.Value || this.Comando.ExecuteScalar() != null)
                //escalar = int.Parse(this.Comando.ExecuteScalar().ToString());
            }
            catch (InvalidCastException ex)
            {
                throw new Exception("Error al ejecutar un escalar.", ex);
            }
            return escalar;
        }
        /// <summary>
        /// Ejecuta el comando creado y retorna un escalar.
        /// </summary>
        /// <returns>El escalar que es el resultado del comando.</returns>
        /// <exception cref="BaseDatosException">Si ocurre un error al ejecutar el comando.</exception>
        public string EjecutarEscalarCadena()
        {
            string escalar = string.Empty;
            object objeto = null;
            try
            {
                objeto = this.Comando.ExecuteScalar();
                if (objeto != null)
                {
                    escalar = objeto.ToString();
                }
            }
            catch (InvalidCastException ex)
            {
                throw new Exception("Error al ejecutar un escalar.", ex);
            }
            return escalar;

            //string escalar = string.Empty;
            //try
            //{
            //    if (this.Comando.ExecuteScalar() != DBNull.Value || this.Comando.ExecuteScalar() != null)
            //        escalar = Convert.ToString(this.Comando.ExecuteScalar());
            //}
            //catch (InvalidCastException ex)
            //{
            //    throw new Exception("Error al ejecutar un escalar.", ex);
            //}
            //return escalar;
        }

        /// <summary>
        /// Ejecuta el comando creado.
        /// </summary>
        public int EjecutarComando()
        {
            return this.Comando.ExecuteNonQuery();
        }

        public void EjecutarComando(string nomPara, ref int valorSalida)
        {
            this.Comando.ExecuteNonQuery();
            valorSalida = Int32.Parse((Comando.Parameters[nomPara].Value.ToString()));
        }

        public void EjecutarComando(string nomPara, ref string valorSalida)
        {
            this.Comando.ExecuteNonQuery();
            valorSalida = Comando.Parameters[nomPara].Value.ToString();
        }
        /// <summary>
        /// Comienza una transacción en base a la conexion abierta.
        /// Lo que se ejecute luego de esta ionvocación estará
        /// dentro de una tranasacción.
        /// </summary>
        public void ComenzarTransaccion()
        {
            if (this.Transaccion == null)
            {
                this.Transaccion = this.Conexion.BeginTransaction();
            }
        }

        public void ComenzarTransaccion2()
        {
            if (this.Transaccion == null)
            {
                this.Transaccion = this.Conexion.BeginTransaction(IsolationLevel.Serializable);
            }
        }

        /// <summary>
        /// Cancela la ejecución de una transacción.
        /// Lo ejecutado entre ésta invocación y su
        /// correspondiente <c>ComenzarTransaccion</c> será perdido.
        /// </summary>
        public void CancelarTransaccion()
        {
            if (this.Transaccion != null)
            {
                this.Transaccion.Rollback();
            }
        }

        /// <summary>
        /// Confirma todo los comandos ejecutados entre el <c>ComanzarTransaccion</c>
        /// y ésta invocación.
        /// </summary>
        public void ConfirmarTransaccion()
        {
            if (this.Transaccion != null)
            {
                this.Transaccion.Commit();
                Transaccion = null;
            }
        }

        /// <summary>
        /// Regresa la consulta en en DataTable
        /// </summary>
        /// <returns></returns>
        public DataTable EjecutarConsultaTabla()
        {
            Adapter = new OleDbDataAdapter(Comando);
            Tabla = new DataTable();
            Adapter.Fill(Tabla);
            return Tabla;
        }

        public OleDbDataAdapter RegresaAdapter()
        {
            Adapter = new OleDbDataAdapter(Comando);
            return Adapter;
        }
    }
}
