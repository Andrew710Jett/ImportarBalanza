using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportarBalanza
{
    public class AUT
    {
        //public int IdCia { get; set; }
        //public string NombreCia { get; set; }
        //public string Tipo { get; set; }
        //public string TipoVehiculo { get; set; }
        //public string Periodo { get; set; }
        ////public string Trimestre { get; set; }
        //public int PmaDirecta { get; set; }
        //public int PorcDirecta { get; set; }
        //public int PmaCedida { get; set; }
        //public int PmaRetenida { get; set; }
        //public int ResRiesgosRetenida { get; set; }
        //public int PmaDevRetenida { get; set; }
        //public int CobExcPerdida { get; set; }
        //public int AdqDirecta { get; set; }
        //public int SinRetenida { get; set; }
        //public int ResTecnico { get; set; }
        //public int ResOperacionAnalog { get; set; }
        //public int GastosOp { get; set; }
        //public int ResOperacion { get; set; }
        //public int ProdFinan { get; set; }
        //public int OtraReserva { get; set; }
        //public int IndCombinado { get; set; }
        //public int CtoNetoAdq { get; set; }

        
        public bool CargarData(DataTable tbData) {


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

                    s.DestinationTableName = "AUTR_FNAMIS";

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

        public void CorreoSP(int num)
        {
            try
            {
                Comando.Connection = Conexion.AbrirConexion();
                Comando.CommandText = "AUTR_FinaciAmisCarga";
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


