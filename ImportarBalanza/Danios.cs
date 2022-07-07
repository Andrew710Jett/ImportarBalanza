using System;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;


namespace ImportarBalanza
{
    public class Danios
    {

        public bool CargarData(DataTable tbData)
        {


            bool resultado = true;
            using (SqlConnection cn = new SqlConnection(Configuracion.Conexion))
            {
                cn.Open();
                using (SqlBulkCopy s = new SqlBulkCopy(cn))
                {
                    s.ColumnMappings.Add("Periodo", "Periodo");
                    s.ColumnMappings.Add("Ramo/Operación", "TipoVehiculo");
                    s.ColumnMappings.Add("tamanio_cia", "Tipo");
                    s.ColumnMappings.Add("#", "Id");
                    s.ColumnMappings.Add("nombre_compania", "NombreCia");
                    s.ColumnMappings.Add("Prima Directa", "PmaDirecta");
                    s.ColumnMappings.Add("Prima Tomada", "PmaTomada");
                    s.ColumnMappings.Add("prima cedida", "PmaCedida");
                    s.ColumnMappings.Add("Prima Retenida", "PmaRetenida");
                    s.ColumnMappings.Add("Incremento Reserva", "ResRiesgosRetenida");
                    s.ColumnMappings.Add("prima devengada", "PmaDevRetenida");
                    s.ColumnMappings.Add("Cob. Exceso Pérdida", "CobExcPerdida");
                    s.ColumnMappings.Add("Costo adquisición", "AdqDirecta");
                    s.ColumnMappings.Add("Costo Neto", "CtoNetoAdq");
                    s.ColumnMappings.Add("Costo siniestralidad", "SinRetenida");
                    s.ColumnMappings.Add("Resultado Técnico", "ResTecnico");
                    s.ColumnMappings.Add("Res. Oper. Análogas", "ResOperacionAnalog");
                    s.ColumnMappings.Add("Gastos Operación", "GastosOp");
                    s.ColumnMappings.Add("Resultado Operación", "ResOperacion");
                    s.ColumnMappings.Add("Producto Financiero", "ProdFinan");
                    s.ColumnMappings.Add("Otras reservas", "OtraReserva");
                    s.ColumnMappings.Add("Res. Inversion Perm.", "ResInversionPerm");
                    s.ColumnMappings.Add("Impuesto Utilidad", "ImpuestoUtilidad");
                    s.ColumnMappings.Add("Resultado Neto", "ResultadoNeto");
                    s.ColumnMappings.Add("indice combinado", "IndCombinado");
                    //s.ColumnMappings.Add("trimestre", "Trimestre");
                    //s.ColumnMappings.Add("id", "IdCia");

                    s.DestinationTableName = "DAN_FNAMIS";

                    s.BulkCopyTimeout = 1500;
                    try
                    {
                        s.WriteToServer(tbData);
                    }
                    catch (Exception e)
                    {
                        string st = e.Message;
                        Console.WriteLine("----------------------->" + st);
                        Trace.WriteLine("----------------------->" + st);
                        resultado = false;
                    }
                }
            }
            return resultado;
        }

        private SqlCommand Comando = new SqlCommand();
        private Conex Conexion = new Conex();

        public void CorreoSPDN(int num)
        {
            try
            {
                Comando.Connection = Conexion.AbrirConexion();
                Comando.CommandText = "DAN_FinaciAmisCarga";
                Comando.CommandType = CommandType.StoredProcedure;
                Comando.Parameters.AddWithValue("@accion", num);
                Comando.ExecuteNonQuery();
                Comando.Parameters.Clear();
                Conexion.CerrarConexion();
            }
            catch (Exception sqlError)
            {
                Trace.WriteLine("Error sql --> " + sqlError);
            }
        }

    }
}
