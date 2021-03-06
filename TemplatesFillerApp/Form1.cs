﻿using System;
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

            htmlManipulator.loadData(textBox1.Text,statusWindow);

            workbookManipulator.loadData(textBox2.Text, statusWindow);
        }



        private String OpenFolderDialog()
        {

            OpenFileDialog folderBrowser = new OpenFileDialog();
            // Set validate names and check file exists to false otherwise windows will
            // not let you select "Folder Selection."
            folderBrowser.ValidateNames = false;
            folderBrowser.CheckFileExists = false;
            folderBrowser.CheckPathExists = true;
            // Always default to Folder Selection.
            folderBrowser.FileName = "Folder Selection.";
            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                return folderBrowser.FileName;
                // ...
            }

            return "";


        }


        private void sendEmails() {

            SmtpClient client = new SmtpClient(comboBox1.Text,Int32.Parse(textBox6.Text));

            client.EnableSsl = true;

            client.Credentials = new System.Net.NetworkCredential(textBox4.Text, textBox5.Text);
            

            foreach (DataRow row in workbookManipulator.excelData.Rows)
            {
               
                MailAddress to = new MailAddress(""+row["Email to"]);
                 
                MailAddress from = new MailAddress(textBox4.Text);

                MailMessage email = new MailMessage(from, to);

                email.Subject = textBox3.Text;

                email.IsBodyHtml = true;

                email.Body = htmlManipulator.createBodyText(workbookManipulator.excelData);

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

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox2.Text = OpenFolderDialog();

        }

        private void folderBrowserDialog1_HelpRequest_1(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Text = OpenFolderDialog();
        }
    }
}
