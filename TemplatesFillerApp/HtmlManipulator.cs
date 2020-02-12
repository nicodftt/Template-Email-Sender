using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;





namespace TemplatesFillerApp
{
    class HtmlManipulator : Manipulator

    {
       private String[] attachments;

       private String bodyText;
        public void loadData(String htmlPath, Form3 statusWindow)
        {

           
            validatePath(htmlPath,statusWindow);

                statusWindow.SetText("Loading Html file...");

             bodyText = File.ReadAllText(htmlPath);
        }

        private void loadAttachments(HtmlNode[] imageNodes) {

            foreach (HtmlNode item in imageNodes)
            {
                attachments.Append(item.GetAttributeValue("src","not found"));
            }
        } 

       private String replaceImageTags(String htmlText)
        {
            HtmlDocument html = new HtmlDocument();
            html.LoadHtml(htmlText);
            HtmlNode[] imageNodes = html.DocumentNode.SelectNodes("\\img").ToArray();
            loadAttachments(imageNodes);

                        

            return "";
        }

         public String createBodyText(System.Data.DataTable excel)
        {


            String auxiliarBody = bodyText;

            foreach (DataColumn column in excel.Columns)
            {
                String columnName = column.ColumnName;

                auxiliarBody = auxiliarBody.Replace("*[" + column.ColumnName + "]", "" + row[column.ColumnName]);

            }

            return auxiliarBody;

        }
    }
}
