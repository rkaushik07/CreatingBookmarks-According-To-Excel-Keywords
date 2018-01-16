using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Diagnostics;
//___ iTextSHarp____<.dll(s) to be found in the dll folder in the Team#12_Q5/dll>
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
//____Aspose____<.dll(s) to be found in the dll folder in the Team#12_Q5/dll>
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using Microsoft.Office.Interop.Excel;

namespace codeFestOwn
{
   
    class Program
    {
        // Folder Path to Program.cs of the solution {Must Be changed By User}
      public static string pathToprogram_cs = @"D:\Code Fest 2.0\Solution_folder_for_Q5\codeFestOwn\";
        // folder Path to pdf Documents
       public static string filePath = pathToprogram_cs + @"\pdf_files\";
        // document names
       public static string[] files = { "Document1.pdf", "Document2.pdf" };
        // file path for Excel worksheet
       public static string excelFilePath = pathToprogram_cs + @"\excelWorkSheet\BookmarkList.xlsx";
        // List of all key and page number - bookmarkNode
        static public IList<bookMarkNode> pageNo = new List<bookMarkNode>();
        
        public static void processStart()
        {
            for (int i = 0; i < 2; i++)
            {
                processPdf.processPDF(Program.files[i], i + 1);
            }
        }
    }
}
