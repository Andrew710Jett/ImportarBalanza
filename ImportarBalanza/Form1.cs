using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Z.Dapper.Plus;

namespace ImportarBalanza
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        int contar = 0;
        public Form1()
        {
            InitializeComponent();
  

            lblCont.Text = contar.ToString();
        }

        
        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = tableCollection[comboBox1.SelectedItem.ToString()];
                if (dt != null)
                {
                    List<Balanza> list = new List<Balanza>();
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        Balanza obj = new Balanza();
                        obj.CodigoCuenta = dt.Rows[i]["CodigoCuenta"].ToString();
                        Console.WriteLine(obj.CodigoCuenta);
                        obj.Cuenta       = dt.Rows[i]["Cuenta"].ToString();
                        Console.WriteLine(obj.Cuenta);
                        obj.Descripcion  = dt.Rows[i]["Descripcion"].ToString();
                        Console.WriteLine(obj.Descripcion);
                        obj.Ramo         = dt.Rows[i]["Ramo"].ToString();
                        Console.WriteLine(obj.Ramo);
                        //obj.BalnceIni    = Convert.ToDouble(dt.Rows[i]["BalnceIni"]);
                        //Console.WriteLine(obj.BalnceIni);
                        //obj.TotalDR      = Convert.ToDouble(dt.Rows[i]["TotalDR"]);
                        //Console.WriteLine(obj.TotalDR);
                        //obj.TotalCR      = Convert.ToDouble(dt.Rows[i]["TotalCR"]);
                        //Console.WriteLine(obj.TotalCR);
                        obj.EndBalance   = Convert.ToDouble(dt.Rows[i]["EndBalance"]);
                        Console.WriteLine(obj.EndBalance);
                        obj.Periodo      = dt.Rows[i]["Periodo"].ToString();
                        Console.WriteLine(obj.Periodo);
                        list.Add(obj);
                    }
                    dataGridView1.DataSource = list;
                }
                else
                {
                    MessageBox.Show("La Hoja No Tiene El Formato Adecuado", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Mensaje", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }

        DataTableCollection tableCollection;
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls" })

                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        lblRuta.Text = ofd.FileName;
                        using (var stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read))
                        {
                            using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                            {
                                DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                                {
                                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                    {
                                        UseHeaderRow = true
                                    }
                                });
                                tableCollection = result.Tables;
                                comboBox1.Items.Clear();
                                foreach (DataTable table in tableCollection)
                                    comboBox1.Items.Add(table.TableName);
                            }
                        }
                    }
            }
            catch (Exception a)
            {

                Trace.WriteLine("Error Archivo Abierto en 2do Plano" + a);

            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }
        public void llamadaProcedure()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection("server = LNMXVMSQL06; database = SS2; integrated security = true"))
                {
                    conn.Open();

                    // 1.  create a command object identifying the stored procedure
                    SqlCommand cmd = new SqlCommand("LimpiarBalanza", conn);

                    // 2. set the command object so it knows to execute a stored procedure
                    cmd.CommandType = CommandType.StoredProcedure;

                    // execute the command
                    using (SqlDataReader rdr = cmd.ExecuteReader())
                    {
                        // iterate through results, printing each to console
                        while (rdr.Read())
                        {
                            Console.WriteLine("Product: {0,-35} Total: {1,2}", rdr["ProductName"], rdr["Total"]);
                        }
                    }
                }
                Trace.WriteLine("Procedimiento Exitoso");
            }
            catch (Exception eo)
            {
                Trace.WriteLine("Error en el procedimiento falló" + eo);
            }

        }

        //     using (SqlConnection conn = new SqlConnection("<connection string>")) { 
        //    conn.Open(); 

        //    string query = "INSERT INTO NombreTabla (campo1, campo2) VALUES (@aram1, @param2)";
        //    SqlCommand cmd = new SqlCommand(query, conn); 


        //    foreach (DataGridViewRow row in dataGridView1.Rows) { 
        //        cmd.Parameters.Clear(); 

        //        cmd.Parameters.AddWithValue("@param1", Convert.ToString(row.Cells["nombreCol1"].Value)); 
        //        cmd.Parameters.AddWithValue("@param2", Convert.ToInt32(row.Cells["nombrecol2"].Value)); 

        //        cmd.ExecuteNonQuery(); 
        //    } 
        //}
        Conex con = new Conex();

        private void button1_Click(object sender, EventArgs e)
        {
             try
            {
                if (lblRuta.Text == null || lblRuta.Text == "")
                {
                    MessageBox.Show("Debes Selecionar el Archivo Correcto de la 'Balanza'","Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                llamadaProcedure();


                    String connectionString = "server = LNMXVMSQL06; database = SS2; integrated security = true";

                    DapperPlusManager.Entity<Balanza>().Table("updatebalanza");
                    List<Balanza> balanza = dataGridView1.DataSource as List<Balanza>;
                    if (balanza != null)
                    {
                        using (IDbConnection db = new SqlConnection(connectionString))
                        {
                            db.BulkInsert(balanza);
                        }
                    }

                    Console.WriteLine(404);
                    MessageBox.Show("Importación Completa!!", "Proceso Correcto", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    lblCont.Text = "0";
                    lblRuta.Text = null;
                    comboBox1.Items.Clear();
                    dataGridView1.DataSource = null;

                }
            }
             catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Console.WriteLine("Error: ---->" + ex);
            }
        }

        private void updateBalanzaBindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            lblCont.Text = dataGridView1.Rows.Count.ToString();
        }
    }
}
     
