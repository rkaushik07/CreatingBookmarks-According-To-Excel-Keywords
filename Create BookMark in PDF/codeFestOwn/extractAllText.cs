using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using Microsoft.Office.Interop.Excel;
using iTextSharp.text.pdf.parser;

namespace codeFestOwn
{

    // read through the text in the PDF and saving as string for future use

    class extractAllText
    {
        public static string ExtractAllTextFromPdf(string inputFilePath)
        {
            //Sanity checks
            if (string.IsNullOrEmpty(inputFilePath))
                throw new ArgumentNullException("inputFile");
            if (!System.IO.File.Exists(inputFilePath))
                throw new System.IO.FileNotFoundException("Cannot find inputFile", inputFilePath);

            //Create a stream reader
            using (FileStream SR = new FileStream(inputFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                //Create a reader to read the PDF
                PdfReader reader = new PdfReader(SR);

                //Create a buffer to store text
                StringBuilder Buf = new StringBuilder();

                //Use the PdfTextExtractor to get all of the text on a page-by-page basis
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    Buf.AppendLine(PdfTextExtractor.GetTextFromPage(reader, i));
                    //Console.WriteLine(PdfTextExtractor.GetTextFromPage(reader, i));
                }
                return Buf.ToString();
            }
        }
    }
}
