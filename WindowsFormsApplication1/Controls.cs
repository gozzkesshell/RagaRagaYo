using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace WindowsFormsApplication1
{
    class Controls
    {

        static public void WriteToExcel()
        {

        }

        static public void BandAnswerCreator()
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


            foreach(object item in bkmrk )
            {


            }
            //object oBookMark = "MyBookmark";
            //doc.Bookmarks.get_Item(ref oBookMark).Range.Text = "Some Text Here";
        }


        private void ReplaceBookmarkText(Microsoft.Office.Interop.Word.Document doc, string bookmarkName, string text)
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


        static public void TicketCreator()
        {

        }
    }

  }

