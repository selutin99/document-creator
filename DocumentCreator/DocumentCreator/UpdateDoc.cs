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

        public void updateDoc(Dictionary<string, Object> keyValuePairs)
        {
            //
            Word.Application wordApp = new Word.Application();
            Word.Document doc = wordApp.Documents.Open(documentPath, ReadOnly: false);
            doc.Activate();
            FindAndReplace(wordApp, "{id:name}", keyValuePairs["{id:name}"]);
            FindAndReplace(wordApp, "{id:theme}", keyValuePairs["{id:theme}"]);
            FindAndReplace(wordApp, "{id:themeName}", keyValuePairs["{id:themeName}"]);
            FindAndReplace(wordApp, "{id:lesson}", keyValuePairs["{id:lesson}"]);
            FindAndReplace(wordApp, "{id:lessonName}", keyValuePairs["{id:lessonName}"]);
            FindAndReplace(wordApp, "{id:kind}", keyValuePairs["{id:kind}"]);
            FindAndReplace(wordApp, "{id:method}", keyValuePairs["{id:method}"]);
            FindAndReplace(wordApp, "{id:duration}", keyValuePairs["{id:duration}"]);
            FindAndReplace(wordApp, "{id:place}", keyValuePairs["{id:place}"]);
            FindAndReplace(wordApp, "{id:literature}", keyValuePairs["{id:literature}"]);
            FindAndReplace(wordApp, "{id:additionalLiterature}", keyValuePairs["{id:additionalLiterature}"]);
            FindAndReplace(wordApp, "{id:technicalMeans}", keyValuePairs["{id:technicalMeans}"]);
            FindAndReplace(wordApp, "{id:intro}", keyValuePairs["{id:intro}"]);
            FindAndReplace(wordApp, "{id:material}", keyValuePairs["{id:material}"]);
            FindAndReplace(wordApp, "{id:educationalQuestions}", keyValuePairs["{id:educationalQuestions}"]);
            FindAndReplace(wordApp, "{id:conclution}", keyValuePairs["{id:conclution}"]);
            foreach(Word.Table table in doc.Tables)
            {
                Word.Range rangeTable = table.Range;
                foreach (Word.Row row in rangeTable.Rows) {
                    foreach (Word.Cell cell in row.Cells)
                    {
                        Word.Range range = cell.Range;
                        if (range.Text.Contains("{id:questions}"))
                        {
                            FindAndReplace(wordApp, "{id:questions}", "");
                            Word.Row temporary=row;
                            int count = 1;
                            foreach (KeyValuePair<string,string> question in (Dictionary<string,string>)keyValuePairs["{id:questions}"])
                            {
                                Object oMissing = System.Reflection.Missing.Value;
                                Word.Row newRow = table.Rows.Add(ref oMissing);
                                newRow.Cells[1].Range.Text = "2."+count;
                                newRow.Cells[2].Range.Text = question.Key;
                                newRow.Cells[3].Range.Text = question.Value;
                                temporary = newRow;
                                count++;
                            }
                            Object missing = System.Reflection.Missing.Value;
                            Word.Row newRowENd = table.Rows.Add(ref missing);
                            newRowENd.Cells[1].Range.Text = "3";
                            newRowENd.Cells[2].Range.Text = "Заключение";
                            newRowENd.Cells[3].Range.Text = (string)keyValuePairs["{id:conclution}"]+ " мин";
                        }
                        else if (range.Text.Contains("{id:goal}"))
                        {
                            FindAndReplace(wordApp, "{id:goal}", "");
                            int count = 1;
                            foreach (string goal in (List<string>)keyValuePairs["{id:goal}"])
                            {
                                Object oMissing = System.Reflection.Missing.Value;
                                Word.Row newRow = table.Rows.Add(ref oMissing);
                                newRow.Cells[1].Range.Text = count+"";
                                newRow.Cells[2].Range.Text = goal;
                                count++;
                            }
                        }
                    }
                }
            }
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


