using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TemplatesFillerApp
{
    class  Manipulator
    {
        public static void validatePath(String path,Form3 statusWindow)
        {

            if (!File.Exists(path))
                throw new Exception("Incorrect path: " + path);
            else
                statusWindow.SetText("Path:" + path + " Ok...");

        }       

    }
}
