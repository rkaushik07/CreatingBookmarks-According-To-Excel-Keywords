using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace codeFestOwn
{
    class countOccurences
    {
        // checking for occureneces of keys in a PDF page and adding in a list
        public static int countNoOfOccurencesOfKey(string input, string toCountString, string fileName)
        {
            // Opening File Strem to browse through PDF and find the keys instances
            using (FileStream SR = new FileStream(Program.filePath + fileName, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                //Create a reader to read the PDF
                PdfReader reader = new PdfReader(SR);

                //Create a buffer to store text
                StringBuilder Buf = new StringBuilder();

                //Use the PdfTextExtractor to get all of the text on a page-by-page basis
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    // Checking for occurences of the key page wise. 
                    if (Regex.Matches(PdfTextExtractor.GetTextFromPage(reader, i), toCountString).Count > 0)
                    {
                        //Storing Values into the Global List
                       Program.pageNo.Add(new bookMarkNode(toCountString, i));
                    }
                }
                // returning number of total occurences of a key in the whole PDF
                return Regex.Matches(input, toCountString).Count;
            }
        }
    }
}
