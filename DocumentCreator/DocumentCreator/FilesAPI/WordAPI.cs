using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Word = Microsoft.Office.Interop.Word;

namespace DocumentCreator.FilesAPI
{
    public class WordAPI
    {
        public static Word.Document GetDocument(string fileName)
        {
            Word.Application app = new Word.Application();
            app.Visible = true;
 
            Word.Document doc = null;

            try
            {
                doc = app.Documents.Open(fileName);
            }
            catch (Exception e)
            {
                throw new Exception("Can't open file", e);
            }
            return doc;
        }

        public static void saveFile(Word.Document doc, string fileName = "")
        {
            if (string.IsNullOrEmpty(fileName))
            {
                try
                {
                    doc.Save();
                }
                catch (Exception e)
                {
                    throw new Exception("Can't save file", e);
                }
            }
            else
            {
                try
                {
                    doc.SaveAs(fileName);
                }
                catch (Exception e)
                {
                    throw new Exception("Can't save file in " + fileName, e);
                }
            }
        }

        public static void close(Word.Document doc)
        {
            if (doc != null)
            {
                doc.Close();
                killWord();
            }
            else
            {
                throw new NullReferenceException();
            }
        }

        public static void killWord()
        {
            System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcessesByName("WORD");
            foreach (System.Diagnostics.Process p in procs)
            {
                p.Kill();
            }
        }
    }
}
