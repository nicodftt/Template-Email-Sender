using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;



namespace TemplatesFillerApp
{
    class HtmlManipulator : Manipulator
    {

       public String loadData(String htmlPath, Form3 statusWindow)
        {

            validatePath(htmlPath,statusWindow);

                statusWindow.SetText("Loading Html file...");

            return File.ReadAllText(htmlPath);
        }

       private String replaceImageTags(String htmlText)
        {
            return "";
        }
    }
}
