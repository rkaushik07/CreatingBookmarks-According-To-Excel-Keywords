using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace codeFestOwn
{
    //complete execution of PDF acording to requirement

    class processPdf
    {
       // static string pathToprogram_cs = @"C:\Users\rkaushik\Source\Workspaces\CodeFest2.0_2017\Team#12_Q5\Source\Solution_folder_for_Q5\codeFestOwn";
        //static string filePath = Program.pathToprogram_cs + @"\pdf_files\";
        public static void processPDF(string fileName, int documentNumber)
        {
            // instantiate program to fetch keys from worksheet

            fetchData programObject = new fetchData();

            // fetch keys for a PDF document from worksheet
            String[] documentKeySet = programObject.fetchKeysFromExcel(documentNumber);

            // Check if pdf path exists
            if (File.Exists(Program.filePath + fileName))
            {
                // instantiate reader class

                PdfReader pdfReader = new PdfReader(Program.filePath + fileName);

                // Extract all text for use from the designated PDF file and process
                string pdfAsString =extractAllText.ExtractAllTextFromPdf(Program.filePath + fileName);

                // Make bookmarks in the PDF
                makeBookMark.makeBookmarksInPdf(documentKeySet, pdfAsString, fileName);

                // Close reader class object
                pdfReader.Close();
            }

        }
    }
}
