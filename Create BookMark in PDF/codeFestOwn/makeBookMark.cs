using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using iTextSharp.text;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace codeFestOwn
{
    class makeBookMark
    {
        public static void makeBookmarksInPdf(string[] documentKeySet, string pdfAsString, string fileName)
        {
            // Using Aspose 3rd party library for bookmark creation
            Aspose.Pdf.Document document = new Aspose.Pdf.Document(Program.filePath + fileName);
            // Initializing list of parent bookmark tabs
           OutlineItemCollection[] parent = new OutlineItemCollection[documentKeySet.Length];

            // Iteration to build sub-sections to partent bookmarks 
            for (int k = 1; k < documentKeySet.Length; k++)
            {
                // getting total occurences of a key in the PDF
                int keyCount = countOccurences.countNoOfOccurencesOfKey(pdfAsString, documentKeySet[k], fileName);
                // creating individual parent tabs
                parent[k] = new OutlineItemCollection(document.Outlines);
                // Making Parents apprear in bold
                parent[k].Bold = true;

                // Setting names of parent bookmark tabs
                parent[k].Title = documentKeySet[k];

                try
                {
                    // count variable for iteration and naming
                    int count = 0;
                    // checking for pagewise bookmark entry
                    foreach (bookMarkNode item in Program.pageNo)
                    {
                        // chekcing if the key in the page is equal to the current key, then do further operations to add bookmarks
                        if (item.key == documentKeySet[k])
                        {
                            // Creating subsection bookmarks
                            OutlineItemCollection child = new OutlineItemCollection(document.Outlines);
                            // Naming Child Bookmarks
                            child.Title = documentKeySet[k] + "_" + (count + 1);
                            // creating the link of the child bookmark to its respective page
                            child.Destination = new GoToAction(item.page_no);
                            // Inrementing counter variable
                            count = count + 1;
                            // Showing child in Italics
                            child.Italic = true;

                            // Connecting Child to parent bookmarks
                            parent[k].Add(child);
                        }
                    }
                    // adding parent bookmarks to the complete PDF document
                    document.Outlines.Add(parent[k]);

                }
                // Catching exceptions and noting in Output window 
                catch (DocumentException dex)
                {
                    Debug.Write(dex.Message);
                }
                catch (IOException ioex)
                {
                    Debug.Write(ioex.Message);
                }
                //Console.WriteLine("" + keyCount);
            }
            // Save the processed PDF File
            document.Save(Program.filePath + fileName.Replace(".pdf", "") + "_processed.pdf");
        }
    }
}
