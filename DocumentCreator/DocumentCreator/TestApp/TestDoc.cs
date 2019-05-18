using DocumentCreator.FilesAPI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace DocumentCreator.TestApp
{
    class TestDoc
    {
        
        public void getAndParseTestDock()
        {

            Word.Document documetn = WordAPI.GetDocument("C://Users//Nikita//source//repos//document-creator//DocumentCreator//DocumentCreator//FilesAPI//Doc.docx");
            documetn.Content.Text+="Hello worldddd!!!!!!";
            WordAPI.saveFile(documetn, "C://Users//Nikita//newDoc.docx");
        }

    }
}
