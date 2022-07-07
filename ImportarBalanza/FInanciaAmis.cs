using ExcelDataReader;
using ImportarBalanza;
using System;
using System.Collections.Generic;
using System.ComponentModel;
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

namespace ImportarAUT
{
    public partial class FInanciaAmis : Form
    {
        private DataSet dtsTablas = new DataSet();
        int bandera = 0;
        public FInanciaAmis()
        {
            InitializeComponent();
            button1.Visible = false;
            button2.Visible = false;
            comboBox1.Enabled = false;

        }
        
        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            dataGridView1.DataSource = dtsTablas.Tables[comboBox1.SelectedIndex];

            //try
            //{
            //    DataTable dt = tableCollectionAUT[comboBox1.SelectedItem.ToString()];
            //    if (dt != null)
            //    {
            //        List<AUT> list = new List<AUT>();
            //        for (int i = 0; i < dt.Rows.Count; i++)
            //        {
            //            AUT obj = new AUT();

            //            obj.Periodo = dt.Rows[i]["Periodo"].ToString();
            //            Console.WriteLine(obj.Periodo);
            //            obj.TipoVehiculo = dt.Rows[i]["Ramo/OperaciÃ³n"].ToString();
            //            Console.WriteLine(obj.TipoVehiculo);
            //            obj.Tipo = dt.Rows[i]["tamanio_cia"].ToString();
            //            Console.WriteLine(obj.Tipo);
            //            obj.IdCia = Convert.ToInt32(dt.Rows[i]["#"]);
            //            Console.WriteLine(obj.IdCia);
            //            obj.NombreCia = dt.Rows[i]["nombre_compania"].ToString();
            //            Console.WriteLine(obj.NombreCia);
            //            obj.PmaDirecta = Convert.ToInt32(dt.Rows[i]["Prima Directa"]);
            //            Console.WriteLine(obj.PmaDirecta);
            //            obj.PmaCedida = Convert.ToInt32(dt.Rows[i]["prima cedida"]);
            //            Console.WriteLine(obj.PmaCedida);
            //            obj.PmaRetenida = Convert.ToInt32(dt.Rows[i]["Prima Retenida"]);
            //            Console.WriteLine(obj.PmaRetenida);
            //            obj.ResRiesgosRetenida = Convert.ToInt32(dt.Rows[i]["Incremento Reserva"]);
            //            Console.WriteLine(obj.ResRiesgosRetenida);
            //            obj.PmaDevRetenida = Convert.ToInt32(dt.Rows[i]["prima devengada"]);
            //            Console.WriteLine(obj.PmaDevRetenida);
            //            obj.CobExcPerdida = Convert.ToInt32(dt.Rows[i]["Cob. Exceso PÃ©rdida"]);
            //            Console.WriteLine(obj.CobExcPerdida);
            //            obj.AdqDirecta = Convert.ToInt32(dt.Rows[i]["Costo adquisiciÃ³n"]);
            //            Console.WriteLine(obj.AdqDirecta);
            //            obj.SinRetenida = Convert.ToInt32(dt.Rows[i]["Costo siniestralidad"]);
            //            Console.WriteLine(obj.SinRetenida);
            //            obj.ResTecnico = Convert.ToInt32(dt.Rows[i]["Resultado TÃ©cnico"]);
            //            Console.WriteLine(obj.ResTecnico);
            //            obj.ResOperacionAnalog = Convert.ToInt32(dt.Rows[i]["Res. Oper. AnÃ¡logas"]);
            //            Console.WriteLine(obj.ResOperacionAnalog);
            //            obj.GastosOp = Convert.ToInt32(dt.Rows[i]["Gastos OperaciÃ³n"]);
            //            Console.WriteLine(obj.GastosOp);
            //            obj.ResOperacion = Convert.ToInt32(dt.Rows[i]["Resultado OperaciÃ³n"]);
            //            Console.WriteLine(obj.ResOperacion);
            //            obj.ProdFinan = Convert.ToInt32(dt.Rows[i]["Producto Financiero"]);
            //            Console.WriteLine(obj.ProdFinan);
            //            obj.OtraReserva = Convert.ToInt32(dt.Rows[i]["Otras reservas"]);
            //            Console.WriteLine(obj.OtraReserva);
            //            obj.IndCombinado = Convert.ToInt32(dt.Rows[i]["indice combinado"]);
            //            Console.WriteLine(obj.IndCombinado);
            //            obj.CtoNetoAdq = Convert.ToInt32(dt.Rows[i]["Costo Neto"]);
            //            Console.WriteLine(obj.CtoNetoAdq);

            //            list.Add(obj);
            //        }
            //        dataGridView1.DataSource = list;
            //    }
            //    else
            //    {
            //        MessageBox.Show("La Hoja No Tiene El Formato Adecuado", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    }
            //}
            //catch (Exception ex)
            //{
            //    Trace.WriteLine("----------> "+ ex);
            //    MessageBox.Show(ex.Message, "Mensaje", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private SqlCommand Comando = new SqlCommand();
        private Conex Conexion = new Conex();
        public void TruncateTablas(int opc)
        {
            try
            {
                Comando.Connection = Conexion.AbrirConexion();
                Comando.CommandText = "TruncarFinanciamis";
                Comando.CommandType = CommandType.StoredProcedure;
                Comando.Parameters.AddWithValue("@opc", opc);
                Comando.ExecuteNonQuery();
                Comando.Parameters.Clear();
                Conexion.CerrarConexion();
            }
            catch (Exception sqlError)
            {
                Trace.WriteLine("Error sql --> " + sqlError);
            }
        }




        DataTableCollection tableCollectionAUT;
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
//                TruncateTablas(bandera);

                OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls";

            if (ofd.ShowDialog()==DialogResult.OK)
            {
                comboBox1.Items.Clear();
                dataGridView1.DataSource = null;

                label1.Text = ofd.FileName;

                    FileStream fsSource = new FileStream(ofd.FileName, FileMode.Open, FileAccess.Read);

                IExcelDataReader reader = ExcelReaderFactory.CreateReader(fsSource);

                dtsTablas = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration() 
                    { 
                        UseHeaderRow=true
                    }
                }); ;


                foreach (DataTable tabla in dtsTablas.Tables)
                {
                    comboBox1.Items.Add(tabla.TableName);
                }
                comboBox1.SelectedIndex = 0;
                reader.Close();

            }
            }
            catch (Exception a)
            {

                MessageBox.Show("El Archivo puede estar abierto en 2do Plano, favor d validar", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Trace.WriteLine("Error Archivo Abierto en 2do Plano" + a);
            }

            //try
            //{
            //    using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls" })

            //        if (ofd.ShowDialog() == DialogResult.OK)
            //        {
            //            label1.Text = ofd.FileName;
            //            using (var stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read))
            //            {
            //                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
            //                {
            //                    DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
            //                    {
            //                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
            //                        {
            //                            UseHeaderRow = true
            //                        }
            //                    });
            //                    tableCollectionAUT = result.Tables;
            //                    comboBox1.Items.Clear();
            //                    foreach (DataTable table in tableCollectionAUT)
            //                        comboBox1.Items.Add(table.TableName);
            //                }
            //            }
            //        }
            //}
            //catch (Exception a)
            //{

            //    Trace.WriteLine("Error Archivo Abierto en 2do Plano" + a);

            //}
        }

        public void Autos()
        {
            AUT a = new AUT();
            DataTable date = (DataTable)(dataGridView1.DataSource);
            bool result = new AUT().CargarData(date);

            if (result)
            {
                MessageBox.Show("Importación Completa!!", "Proceso Correcto", MessageBoxButtons.OK, MessageBoxIcon.Information);
                label1.Text = null;
                comboBox1.DataSource = null;
                comboBox1.Items.Clear();
                dataGridView1.DataSource = null;
                radioButton1.Checked = false;
                a.CorreoSP(1);
            }
            else
            {
                MessageBox.Show("Hubo un problema", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        public void Danios()
        {
            Danios d = new Danios();
            DataTable date = (DataTable)(dataGridView1.DataSource);
            bool result = new Danios().CargarData(date);

            if (result)
            {
                MessageBox.Show("Importación Completa!!", "Proceso Correcto", MessageBoxButtons.OK, MessageBoxIcon.Information);
                label1.Text = null;
                comboBox1.DataSource = null;
                comboBox1.Items.Clear();
                dataGridView1.DataSource = null;
                radioButton2.Checked = false;
                d.CorreoSPDN(1);
            }
            else
            {
                MessageBox.Show("Hubo un problema", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {                
                if (label1.Text == null || label1.Text == "")
                {
                    MessageBox.Show("Debes Selecionar el Archivo Correcto", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    button2.Enabled = false;
                }
                else
                {
                    if (bandera==1)
                    {
                        Autos();
                    }
                    else if (bandera == 2)
                    {
                        Danios();
                    }
                    else
                    {
                        MessageBox.Show("Favor de seleccionar una opción", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Console.WriteLine("Error: ---->" + ex);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //label1.Text = dataGridView1.Rows.Count.ToString();
        }

        private void FInanciaAmis_Load(object sender, EventArgs e)
        {
            // TODO: esta línea de código carga datos en la tabla 'sS2DataSet.AUTR_FNAMIS' Puede moverla o quitarla según sea necesario.
            //this.aUTR_FNAMISTableAdapter.Fill(this.sS2DataSet.AUTR_FNAMIS);

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                button1.Visible = true;
                button2.Visible = true;
                comboBox1.Enabled = true;
                //                        Autos();
                bandera = 1;
                TruncateTablas(bandera);
                Console.WriteLine("Autos");
            }
            else
            {
                bandera = 0;
                radioButton1.Checked = false;
                Console.WriteLine("NADA");
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                button1.Visible = true;
                button2.Visible = true;
                comboBox1.Enabled = true;
                //Danios();
                bandera = 2;
                TruncateTablas(bandera);
                Console.WriteLine("Danios");
            }
            else
            {
                bandera = 0;
                radioButton2.Checked = false;
                Console.WriteLine("NADA");
            }
        }

        private void radioButton2_MouseClick(object sender, MouseEventArgs e)
        {
            //Danios d = new Danios();
            //d.TruncateTablas(bandera);
        }
    }
}
