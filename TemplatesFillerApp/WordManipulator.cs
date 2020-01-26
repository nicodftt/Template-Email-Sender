using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;



namespace TemplatesFillerApp
{
    class WordManipulator : Manipulator
    {
        private Document doc;

        private Application wordFile;

        private Object file;

        private void closeApplication(Document doc, Application application) {

            doc.Close(false);
            application.Quit(false);
            
        }

       public String loadData(String wordPath, Form3 statusWindow)
        {

            validatePath(wordPath,statusWindow);

            try
            {
                statusWindow.SetText("Loading Word file...");

                wordFile = new Application();

                file = wordPath;

                String textvalue;

                doc = wordFile.Documents.Open(ref file);

                doc.ActiveWindow.Selection.WholeStory();

                doc.ActiveWindow.Selection.Copy();

                textvalue = doc.Content.Text;

                closeApplication(doc, wordFile);

                System.Windows.Forms.Clipboard.Clear();

                statusWindow.SetText("Word file loaded...");

                return "<p style=\"color: red; font - size:40px;\"><pre>" + textvalue+ "<pre><p>";

            }

            catch (Exception message) {

                closeApplication(doc, wordFile);

                throw new Exception(message.Message);
            }
        }
    }
}
