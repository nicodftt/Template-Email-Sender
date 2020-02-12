using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace TemplatesFillerApp
{
    class WorkbookManipulator : Manipulator
    {

        public System.Data.DataTable excelData = new System.Data.DataTable();

        private Workbook workbook;

        private Worksheet worksheet;

        private System.Data.DataTable excelDataTable = new System.Data.DataTable();


        private void copyHeader(Range usedRange) {
            /*Copy the header to the data table
             I think the best practice is generate an object with the static behavior*/
            int columnIndex = 1;

            foreach (Range column in usedRange.Columns) {
               /*I Should review this part*/
                excelDataTable.Columns.Add(column.Rows.Cells[1,columnIndex].Value2);
             
            }          
        
        }
        /*This method must be reviewed*/
        private void copyValues(Range UsedRange) {
            /* This method copy all the data from the rows to the dataTable.
             * I don't know how to do it without using the index counters. I have to investigate and clean this part*/
            int rowIndex = 2;

            int dataTableRowIndex = 0;

            int columnIndex = 1;         

            while(dataTableRowIndex < UsedRange.Rows.Count - 1) {

                excelDataTable.Rows.Add();

                foreach (System.Data.DataColumn column in excelDataTable.Columns)
                {

                    excelDataTable.Rows[dataTableRowIndex][column.ColumnName] = UsedRange.Rows.Cells[rowIndex, columnIndex].Value2;

                    columnIndex += 1;
                }

                columnIndex = 1;

                dataTableRowIndex += 1;

                rowIndex += 1;
                  
            }
        }


        public void loadData(String excelPath, Form3 statusWindow)
        {

          
            validatePath(excelPath, statusWindow);

            try
            {
                statusWindow.SetText("Loading excel file...");

                Object file = excelPath;

                workbook = new Application().Workbooks.Open(excelPath);

                worksheet = workbook.Worksheets[1];

                Range usedRange = worksheet.UsedRange;

                copyHeader(usedRange);

                copyValues(usedRange);

                workbook.Close();

                statusWindow.SetText("Excel file loaded...");

                statusWindow.SetText("Data table Rows: " + excelDataTable.Rows.Count);

                excelData =  excelDataTable;

            }
            catch (Exception message)
            {

                workbook.Close();

                throw new Exception(message.Message);

            }

        }

    }
}
