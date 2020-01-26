using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TemplatesFillerApp
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        public void SetText(String statusMessage) => richTextBox1.Text = richTextBox1.Text + statusMessage + "\n";

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
