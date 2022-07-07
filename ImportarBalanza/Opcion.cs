using ImportarAUT;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ImportarBalanza
{
    public partial class Opcion : Form
    {
        public Opcion()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FInanciaAmis a = new FInanciaAmis();
            a.Show();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 b = new Form1();
            b.Show();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
