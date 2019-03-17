using DocumentCreator.FilesAPI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace DocumentCreator
{
    class ParseWorkPrograming
    {
        private Word.Document doc;
        private Word.Tables table;

        public ParseWorkPrograming(string inputFilePath)
        {
            this.doc = FilesAPI.WordAPI.GetDocument(inputFilePath);
            this.table = doc.Tables;
        }

        public List<string> ParsePlan()
        {
            for(int j = 1; j < table.Count; j++) { 
                Word.Range range = table[j].Range;
                Word.Cells cells = range.Cells;
                List<string> requirementsForStudent = new List<string>();
                if (cells[2].Range.Text.StartsWith("Перечень планируемых"))
                {
                    for (int i = 1; i < cells.Count; i++)
                    {
                        Word.Cell cell = cells[i];
                        Word.Range updateRange = cell.Range;
                        string text = updateRange.Text;
                        if (text.StartsWith("В результате "))
                        {
                            text = text.Substring(text.IndexOf("знать:"));
                            string[] requirements = text.Split(';');
                            foreach(string str in requirements)
                            {
                                requirementsForStudent.Add(str.Replace("\r", ""));
                            }
                        }

                    }
                    return requirementsForStudent;
                }
            }
            WordAPI.Close(doc);
            return null;
        }


    }
}
