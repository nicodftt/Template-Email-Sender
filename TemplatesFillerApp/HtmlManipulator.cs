using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using HtmlAgilityPack;





namespace TemplatesFillerApp
{
    class HtmlManipulator : Manipulator
    {
       public String[] attachments;
       public String loadData(String htmlPath, Form3 statusWindow)
        {

           
            validatePath(htmlPath,statusWindow);

                statusWindow.SetText("Loading Html file...");

            return File.ReadAllText(htmlPath);
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
    }
}
