using DocumentCreator.FilesAPI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace DocumentCreator
{
    class UpdateDoc
    {
        string documentPath;
        public UpdateDoc(string path)
        {
            this.documentPath = path;
        }

        public void updateDoc()
        {
            //
            Word.Application wordApp = new Word.Application();
            Word.Document doc = wordApp.Documents.Open(documentPath, ReadOnly: false);
            doc.Activate();
            FindAndReplace(wordApp, "{id:name}", "НАЗВАНИЕ ДИСЦИПЛИНЫ");
 //           FindAndReplace(wordApp, "{id:name}", "НАЗВАНИЕ ДИСЦИПЛИНЫ");
 //           FindAndReplace(wordApp, "{id:name}", "НАЗВАНИЕ ДИСЦИПЛИНЫ");
 //           FindAndReplace(wordApp, "{id:name}", "НАЗВАНИЕ ДИСЦИПЛИНЫ");
 //           FindAndReplace(wordApp, "{id:name}", "НАЗВАНИЕ ДИСЦИПЛИНЫ");
  //          FindAndReplace(wordApp, "{id:name}", "НАЗВАНИЕ ДИСЦИПЛИНЫ");
 //           FindAndReplace(wordApp, "{id:name}", "НАЗВАНИЕ ДИСЦИПЛИНЫ");
 //           FindAndReplace(wordApp, "{id:name}", "НАЗВАНИЕ ДИСЦИПЛИНЫ");
 //           FindAndReplace(wordApp, "{id:name}", "НАЗВАНИЕ ДИСЦИПЛИНЫ");
 //           FindAndReplace(wordApp, "{id:name}", "НАЗВАНИЕ ДИСЦИПЛИНЫ");
            WordAPI.SaveFile(doc);
            WordAPI.Close(doc);
        }

        private void FindAndReplace(Word.Application doc, object findText, object replaceWithText)
        {
            //options
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            //execute find and replace
            doc.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }
    }
}


