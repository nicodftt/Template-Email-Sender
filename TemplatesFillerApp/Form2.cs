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
    public partial class Form2 : Form
    {
        Form1 objectSender;
        public Form2()
        {
            InitializeComponent();
        }

        public void setExceptionMessage(Form1 sender,String exceptionMessage)
        {
            objectSender = sender;
            objectSender.Enabled = false;
            richTextBox1.Text = exceptionMessage;
            richTextBox1.Enabled = false;
        }
        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            objectSender.Enabled = true;
            this.Close(); 
        }
    }
}
