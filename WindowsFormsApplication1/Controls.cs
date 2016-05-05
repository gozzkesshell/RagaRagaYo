using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplication1
{
    static class Controls
    {

        //static public void WriteToExcel()
        //{

        //}

        static public void TicketCreator(Microsoft.Office.Interop.Excel.Worksheet sh1)
        {

            var Word = new Microsoft.Office.Interop.Word.Application();
            object fileName = "C:\\Users\\Gleb Naymitenko\\Documents\\WisitorTicket.docx";
            object readOnly = false;
            object isVisible = true;
            object missing = System.Reflection.Missing.Value;


            Document doc = Word.Documents.Open(ref fileName, ref missing, ref readOnly, ref readOnly,
                           ref missing, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing,
                           ref missing, ref missing, ref missing, ref missing, ref missing);
            String [] bkmrk = { "WisitorName","FestName","WisitorNumber","FestUrl" };

            //getting number of last used row
            int lastRow = sh1.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            //Important data begins from 6 row
            int fromRow = 6;

            //Writing data from excel to word bookmarks
            //E
            while(fromRow <= lastRow)
            {
                fromRow++;
                foreach (string item in bkmrk)
                {

                    ReplaceBookmarkText(doc, item, sh1.Cells[fromRow, "B"]);
                    ReplaceBookmarkText(doc, item, "Fest");
                    ReplaceBookmarkText(doc, item, sh1.Cells[fromRow, "A"]);
                    ReplaceBookmarkText(doc, item, "someadress");

                    doc.SaveAs2((string)sh1.Cells[fromRow, "B"].Replace(" ", ""));
                }
            }

               

          
        }


        static public void ReplaceBookmarkText(Microsoft.Office.Interop.Word.Document doc, string bookmarkName, string text)
        {

            if (doc.Bookmarks.Exists(bookmarkName))

            {

                Object name = bookmarkName;

                Microsoft.Office.Interop.Word.Range range =  doc.Bookmarks.get_Item(ref name).Range;

                range.Text = text;

                object newRange = range;

                doc.Bookmarks.Add(bookmarkName, ref newRange);

            }

        }


        //static public void bandResponseCreator()
        //{

        //}
    }

  }

