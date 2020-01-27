using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Mail;
using Microsoft.Office.Interop.Excel;




namespace TemplatesFillerApp
{
    public partial class Form1 : Form
    {
        private Form3 statusWindow;

        private String bodyText;

        private System.Data.DataTable excelData = new System.Data.DataTable();

        HtmlManipulator htmlManipulator = new HtmlManipulator();

        WorkbookManipulator workbookManipulator = new WorkbookManipulator();

        public Form1()
        {
            InitializeComponent();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void openErrorWindows(String exception)
        {

            Form2 popUp = new Form2();

            popUp.setExceptionMessage(this, exception);

            popUp.Show();

        }

        private void getInfoFromFiles()
        {

            statusWindow.SetText("Loading files...");

            bodyText = htmlManipulator.loadData(textBox1.Text,statusWindow);

            excelData = workbookManipulator.loadData(textBox2.Text, statusWindow);
        }

        private String createBodyText(DataRow row)
        {

           String auxiliarBody = bodyText;

            foreach (DataColumn column in excelData.Columns)
            {
                String columnName=column.ColumnName;

               auxiliarBody = auxiliarBody.Replace("*["+column.ColumnName+"]",""+row[column.ColumnName]);

            }
            return auxiliarBody;
        
        }

        private void sendEmails() {

            SmtpClient client = new SmtpClient("smtp.gmail.com",587);

            client.EnableSsl = true;

            client.Credentials = new System.Net.NetworkCredential(textBox4.Text, textBox5.Text);


            foreach (DataRow row in excelData.Rows)
            {
               
                MailAddress to = new MailAddress(""+row["Email to"]);
                 
                MailAddress from = new MailAddress(textBox4.Text);

                MailMessage email = new MailMessage(from, to);

                email.Subject = textBox3.Text;

                email.IsBodyHtml = true;

                email.Body = @createBodyText(row);

                client.Send(email);
            }
         
        }



        private void button1_Click(object sender, EventArgs e)
        {
            try
            {

                statusWindow = new Form3();

                statusWindow.Show();

                getInfoFromFiles();

                sendEmails();

                statusWindow.SetText("Process completed ...");


            }
            catch (Exception exceptionMessage)
            {

                String statusBar = statusWindow.Text;
                statusWindow.Close();
                openErrorWindows(statusBar + exceptionMessage.Message);

            }
        }

        private void label3_Click_1(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
