using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportarBalanza
{
    public class Conex
    {
        static private string CadenaConexion = "server = LNMXVMSQL06; database = SS2; integrated security = true";
        private SqlConnection Conexion = new SqlConnection(CadenaConexion);
        public SqlConnection AbrirConexion()
        {
            try
            {
                if (Conexion.State == ConnectionState.Closed)
                    Conexion.Open();
                return Conexion;
            }
            catch (SqlException sql)
            {
                throw sql;
            }

        }
        public SqlConnection CerrarConexion()
        {
            try
            {
                if (Conexion.State == ConnectionState.Open)
                    Conexion.Close();
                return Conexion;
            }
            catch (Exception)
            {

                throw;
            }

        }

    }
}
