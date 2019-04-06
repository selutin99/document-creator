using System;
using Word = Microsoft.Office.Interop.Word;

namespace DocumentCreator.FilesAPI
{
    public class WordAPI
    {

        private static Word.Application app = new Word.Application();

        public static Word.Document GetDocument(string fileName)
        {
            //app.Visible = true;
            app = new Word.Application();

            Word.Document doc = null;

            try
            {
                doc = app.Documents.Open(fileName,ReadOnly:false);
            }
            catch (Exception e)
            {
                throw new Exception("Can't open file", e);
            }
            return doc;
        }

        public static Word.Application GetWordApp()
        {
            return app;
        }

        public static void SaveFile(Word.Document doc, string fileName = "")
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

        public static void Close(Word.Document doc)
        {
            if (doc != null)
            {
                doc.Close();
                KillWord();
            }
            else
            {
                throw new NullReferenceException();
            }
        }

        public static void KillWord()
        {
            System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcessesByName("winword");
            foreach (System.Diagnostics.Process p in procs)
            {
                p.Kill();
            }
        }

        public static void FindAndReplace(Word.Document doc, object findText, object replaceWithText)
        {
            doc.Content.Find.Execute(findText, false, true, false, false, false,
                                     true, 1, false, replaceWithText, 2,
                                     false, false, false, false);
        }
    }
}
