 using DocumentCreator.FilesAPI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace DocumentCreator
{
    class UpdateDoc
    {
        private string documentPath;
        private string documentPathPlan;
        char separator = '$';
        public UpdateDoc(string path,string documentPathPlan)
        {
            this.documentPath = path;
            this.documentPathPlan = documentPathPlan;
        }

        public void updateDoc(Dictionary<string, Object> keyValuePairs)
        {
            
            Word.Application wordApp = new Word.Application();
            Word.Document doc = wordApp.Documents.Open(documentPath, ReadOnly: false);
            try { 
            doc.Activate();
            if (((string)(keyValuePairs["{id:kind}"])).Contains("Тренировк")|| ((string)(keyValuePairs["{id:kind}"])).Contains("Самостояте")|| ((string)(keyValuePairs["{id:kind}"])).Contains("Практическо"))
            {
                int countQuestions = 0;
                for (int e=1; e <= doc.Paragraphs.Count; e++)
                {
                    if (doc.Paragraphs[e].Range.Text.Contains("{id:adjunct}"))
                    {
                        Dictionary<string, string> questions = (Dictionary<string, string>)(keyValuePairs["{id:questions}"]);
                        FindAndReplace(wordApp, "{id:adjunct}", "ПРИЛОЖЕНИЯ");
                        countQuestions = questions.Count;
                        Word.Paragraph p= doc.Paragraphs[e];
                        p.Range.InsertParagraphAfter();
                        p = p.Next();
                        for (int i = 0; i < questions.Count; i++)
                        {

                            p.Range.Text= (i + 1)+". Карточка - задание на изучение " + (i + 1) + "-го вопроса.";
                            p.Range.InsertParagraphAfter();
                            p = p.Next();
                        }

                    }
                    int q = 1;
                    int countTables = doc.Tables.Count;
                    if (doc.Paragraphs[e].Range.Text.Contains("{id:file"))
                    {
                        Word.Paragraph p = doc.Paragraphs[e];
                        p.Range.InsertParagraphAfter();
                        p.Next();
                        Dictionary<string, string> questions = (Dictionary<string, string>)(keyValuePairs["{id:questions}"]);
                        int temp = Int32.Parse(doc.Paragraphs[e].Range.Text.Trim().Replace("\r", "").Replace("\a", "").Replace("\n", "").Substring(8, 1));
                        KeyValuePair<string, string> question;
                        try
                        {
                            question = questions.ElementAt(temp - 1);
                        }
                        catch (Exception exception)
                        {
                            break;
                        }
                        p.Range.InsertFile(Path.GetFullPath(Path.Combine(System.Reflection.Assembly.GetExecutingAssembly().Location, @"../../../../../Resources/AdditionalDocs/Spravochnik.docx")));
                        FindAndReplace(wordApp, "{id:questionName}", question.Key);
                        FindAndReplace(wordApp, "{id:countAjunct}", temp);
                        FindAndReplace(wordApp, "{id:questionDuration}", question.Value.Split(separator)[0]);
                    }
                    int afterInsert = doc.Tables.Count;

                }
            }
            if (((string)(keyValuePairs["{id:kind}"])).ToLower().Contains("полувз")|| ((string)(keyValuePairs["{id:kind}"])).ToLower().Contains("ПРО 3 ЧЕЛОВЕК СПРОСИТЬ КАК НАЗЫВАЕТСЯ"))
            {
                foreach (Word.Paragraph paragraph in doc.Paragraphs)
                {
                    if (paragraph.Range.Text.Contains("{id:image}"))
                    {
                        if(((string)(keyValuePairs["{id:kind}"])).ToLower().Contains("ПРО 3 ЧЕЛОВЕК СПРОСИТЬ КАК НАЗЫВАЕТСЯ"))
                        {
                            paragraph.Range.InsertFile(Path.GetFullPath(Path.Combine(System.Reflection.Assembly.GetExecutingAssembly().Location, @"../../../../../Resources/AdditionalDocs/image2.docx")));
                        }
                        else
                        {
                            paragraph.Range.InsertFile(Path.GetFullPath(Path.Combine(System.Reflection.Assembly.GetExecutingAssembly().Location, @"../../../../../Resources/AdditionalDocs/image1.docx")));
                        }
                        
                    }
                }
            }
            FindAndReplace(wordApp, "{id:cardOfTask}", "");
            FindAndReplace(wordApp, "{id:file1}", "");
            FindAndReplace(wordApp, "{id:file2}", "");
            FindAndReplace(wordApp, "{id:file3}", "");
            FindAndReplace(wordApp, "{id:file4}", "");
            FindAndReplace(wordApp, "{id:file5}", "");
            FindAndReplace(wordApp, "{id:file6}", "");
            FindAndReplace(wordApp, "{id:image}", "");
            FindAndReplace(wordApp, "{id:name}", keyValuePairs["{id:name}"]);
            FindAndReplace(wordApp, "{id:theme}", keyValuePairs["{id:theme}"]);
            FindAndReplace(wordApp, "{id:themeName}", keyValuePairs["{id:themeName}"]);
            FindAndReplace(wordApp, "{id:lesson}", keyValuePairs["{id:lesson}"]);
            FindAndReplace(wordApp, "{id:lessonName}", keyValuePairs["{id:lessonName}"]);
            FindAndReplace(wordApp, "{id:kind}", keyValuePairs["{id:kind}"]);
            FindAndReplace(wordApp, "{id:method}", keyValuePairs["{id:method}"]);
            FindAndReplace(wordApp, "{id:duration}", keyValuePairs["{id:duration}"]);
            FindAndReplace(wordApp, "{id:place}", keyValuePairs["{id:place}"]);
            string methodical = (string)keyValuePairs["{id:methodical}"];
            int k = 0;
            for(int i = 0; i < methodical.Length; i += 30)
            {
                if (i > 0)
                {
                    try
                    {
                        FindAndReplace(wordApp, "{id:methodical}", methodical.Substring(i + 1, 30) + "{id:methodical}");
                    }
                    catch (Exception e)
                    {
                        break;
                    }
                }
                else
                {
                    FindAndReplace(wordApp, "{id:methodical}", methodical.Substring(i, 30) + "{id:methodical}");
                }
                    k = i;
            }
            k += 30;
            FindAndReplace(wordApp, "{id:methodical}", methodical.Substring(k + 1));
            string literature = (string)keyValuePairs["{id:literature}"];
            int l = 0;
            for (int i = 0; i < literature.Length; i += 30)
            {
                if (i > 0)
                {
                    try
                    {
                        FindAndReplace(wordApp, "{id:literature}", literature.Substring(i + 1, 30) + "{id:literature}");
                    }
                    catch (Exception e)
                    {
                        break;
                    }
                    
                }
                else
                {
                    FindAndReplace(wordApp, "{id:literature}", literature.Substring(i, 30) + "{id:literature}");
                }
                l = i;
            }
            l += 30;
            FindAndReplace(wordApp, "{id:literature}", literature.Substring(l + 1));
            FindAndReplace(wordApp, "{id:technicalMeans}", keyValuePairs["{id:technicalMeans}"]);
            string introTime = (keyValuePairs["{id:intro}"]).ToString().Split(separator)[0];
            FindAndReplace(wordApp, "{id:intro}", introTime);
            string[] introQuestions = (keyValuePairs["{id:intro}"]).ToString().Split(separator)[1].Split(';');
            for (int e = 0; e < introQuestions.Length - 1; e++)
            {
                FindAndReplace(wordApp, "{id:introQuestions}", introQuestions[e]+ ";\r\n{id:introQuestions}");
            }
            FindAndReplace(wordApp, "{id:introQuestions}","");
            FindAndReplace(wordApp, "{id:material}", keyValuePairs["{id:material}"]);
            FindAndReplace(wordApp, "{id:educationalQuestions}", keyValuePairs["{id:educationalQuestions}"]);
            string conclusionTime = (keyValuePairs["{id:conclution}"]).ToString().Split(separator)[0];
            FindAndReplace(wordApp, "{id:conclution}", conclusionTime);
            string[] conclusionQuestions = (keyValuePairs["{id:conclution}"]).ToString().Split(separator)[1].Split(';');
            for (int e = 0; e < conclusionQuestions.Length - 1; e++)
            {
                FindAndReplace(wordApp, "{id:conclutionsQuestions}", conclusionQuestions[e] + ";\r\n{id:conclutionsQuestions}");
            }
            FindAndReplace(wordApp, "{id:conclutionsQuestions}", "");
            foreach (Word.Table table in doc.Tables)
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
                                newRow.Cells[3].Range.Text = question.Value.Split(separator)[0];
                                Regex regex = new Regex("^ ?[1-9].*$");
                                string questionFull = "";
                                if (regex.IsMatch(question.Key))
                                {
                                    questionFull = "Учебный вопрос. " + question.Key+" "+ question.Value.Split(separator)[0]+ ".\r\n";
                                }
                                else { 
                                    questionFull = "Учебный вопрос " + count + ". " + question.Key + " " + question.Value.Split(separator)[0] + ".\r\n";
                                }
                                int r = 0;
                                for (int e = 0; e < questionFull.Length; e += 30)
                                {
                                    if (e > 0)
                                    {
                                        try
                                        {
                                            FindAndReplace(wordApp, "{id:questionOfLesson}", questionFull.Substring(e, 30) + "{id:questionOfLesson}");
                                        }
                                        catch (Exception q)
                                        {
                                            break;
                                        }

                                    }
                                    else
                                    {
                                        FindAndReplace(wordApp, "{id:questionOfLesson}", questionFull.Substring(e, 30) + "{id:questionOfLesson}");
                                    }
                                    r = e;
                                }
                                r += 30;
                                FindAndReplace(wordApp, "{id:questionOfLesson}", questionFull.Substring(r) + "\r\n{id:contentOfQuestion}");
                                int w = 0;
                                string temp = question.Value.Split(separator)[1].Replace("\n", "\r\n");
                                for (int z=0;z< temp.Length; z+=30)
                                {
                                    if (z > 0)
                                    {
                                        try
                                        {
                                            FindAndReplace(wordApp, "{id:contentOfQuestion}", temp.Substring(z, 30) + "{id:contentOfQuestion}");
                                        }
                                        catch (Exception q)
                                        {
                                            break;
                                        }

                                    }
                                    else
                                    {
                                        FindAndReplace(wordApp, "{id:contentOfQuestion}", temp.Substring(z, 30) + "{id:contentOfQuestion}");
                                    }
                                    w = z;
                                }
                                FindAndReplace(wordApp, "{id:contentOfQuestion}", temp.Substring(w) + "\r\n{id:questionOfLesson}\r\n");
                                temporary = newRow;
                                count++;
                            }
                            Object missing = System.Reflection.Missing.Value;
                            Word.Row newRowENd = table.Rows.Add(ref missing);
                            newRowENd.Cells[1].Range.Text = "3";
                            newRowENd.Cells[2].Range.Text = "Заключение";
                            newRowENd.Cells[3].Range.Text = (string)keyValuePairs["{id:conclution}"].ToString().Split(separator)[0]+ " мин";
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
            FindAndReplace(wordApp, "{id:adjunct}", "");
            FindAndReplace(wordApp, "{id:questionOfLesson}", "");
            WordAPI.SaveFile(doc);
            WordAPI.Close(doc);
            updatePlan(keyValuePairs);
            }
            catch (Exception e)
            {
                doc.Close();
                new ExceptionWindow()
                    .Show();
            }
            
        }










        private void updatePlan(Dictionary<string, Object> keyValuePairs)
        {
            //
            Word.Application wordApp = new Word.Application();
            Word.Document doc = wordApp.Documents.Open(documentPathPlan, ReadOnly: false);
            try { 
            doc.Activate();
            if (((string)(keyValuePairs["{id:kind}"])).Contains("Тренировк") || ((string)(keyValuePairs["{id:kind}"])).Contains("Самостояте") || ((string)(keyValuePairs["{id:kind}"])).Contains("Практическо"))
            {
                int countQuestions = 0;
                Dictionary<string, string> questions = (Dictionary<string, string>)(keyValuePairs["{id:questions}"]);
                for (int e = 1; e <= doc.Paragraphs.Count; e++)
                {
                    if (doc.Paragraphs[e].Range.Text.Contains("{id:adjunct}"))
                    {
                        FindAndReplace(wordApp, "{id:adjunct}", "ПРИЛОЖЕННИЯ");
                        Word.Paragraph p = doc.Paragraphs[e];
                        p.Range.InsertParagraphAfter();
                        p = doc.Paragraphs.Add(p.Range);
                        for (int i = 0; i < questions.Count; i++)
                        {
                            int a = i + 1;
                            p.Range.Text = "Карточка - задание на изучение " + a + "-го вопроса.^l";
                            p.Range.InsertParagraphAfter();
                            p = doc.Paragraphs.Add(p.Range);

                        }
                        for (int i = 0; i < questions.Count; i++)
                        {
                            int counter = doc.Paragraphs.Count - 5+i;
                            doc.Paragraphs[counter].Range.InsertFile(Path.GetFullPath(Path.Combine(System.Reflection.Assembly.GetExecutingAssembly().Location, @"../../../../../Resources/AdditionalDocs/Spravochnik.docx")));
                            KeyValuePair<string, string> question = questions.ElementAt(i);
                            FindAndReplace(wordApp, "{id:questionName}", question.Key);
                            FindAndReplace(wordApp, "{id:countAjunct}", i + 1);
                            FindAndReplace(wordApp, "{id:questionDuration}", question.Value.Split(separator)[0]);
                            if (questions.Count > 1)
                            {
                                countQuestions = 1;
                                break;
                            }

                        }

                    }
                    //if(doc.Paragraphs[e].Range.Text.Trim().Replace("\r","").Replace("\a","").Replace("\n","").Equals("")&&countQuestions>0&&(countQuestions< questions.Count))
                    //{
                    //    doc.Paragraphs[e].Range.InsertFile(Path.GetFullPath(Path.Combine(System.Reflection.Assembly.GetExecutingAssembly().Location, @"../../../../../Resources/AdditionalDocs/Spravochnik.docx")));
                    //    KeyValuePair<string, string> question = questions.ElementAt(countQuestions);
                    //    FindAndReplace(wordApp, "{id:questionName}", question.Key);
                    //    FindAndReplace(wordApp, "{id:countAjunct}", countQuestions + 1);
                    //    FindAndReplace(wordApp, "{id:questionDuration}", question.Value);
                    //    countQuestions++;
                    //}
                    //int q = 1;
                }
            }
            if (((string)(keyValuePairs["{id:kind}"])).ToLower().Contains("полувз") || ((string)(keyValuePairs["{id:kind}"])).ToLower().Contains("ПРО 3 ЧЕЛОВЕК СПРОСИТЬ КАК НАЗЫВАЕТСЯ"))
            {
                foreach (Word.Paragraph paragraph in doc.Paragraphs)
                {
                    if (paragraph.Range.Text.Contains("{id:image}"))
                    {
                        if (((string)(keyValuePairs["{id:kind}"])).ToLower().Contains("ПРО 3 ЧЕЛОВЕК СПРОСИТЬ КАК НАЗЫВАЕТСЯ"))
                        {
                            paragraph.Range.InsertFile(Path.GetFullPath(Path.Combine(System.Reflection.Assembly.GetExecutingAssembly().Location, @"../../../../../Resources/AdditionalDocs/image2.docx")));
                        }
                        else
                        {
                            paragraph.Range.InsertFile(Path.GetFullPath(Path.Combine(System.Reflection.Assembly.GetExecutingAssembly().Location, @"../../../../../Resources/AdditionalDocs/image1.docx")));
                        }

                    }
                }
            }

            FindAndReplace(wordApp, "{id:image}", "");
            FindAndReplace(wordApp, "{id:adjunct}", "");
            FindAndReplace(wordApp, "{id:name}", keyValuePairs["{id:name}"]);
            FindAndReplace(wordApp, "{id:theme}", keyValuePairs["{id:theme}"]);
            FindAndReplace(wordApp, "{id:themeName}", keyValuePairs["{id:themeName}"]);
            FindAndReplace(wordApp, "{id:lesson}", keyValuePairs["{id:lesson}"]);
            FindAndReplace(wordApp, "{id:lessonName}", keyValuePairs["{id:lessonName}"]);
            FindAndReplace(wordApp, "{id:kind}", keyValuePairs["{id:kind}"]);
            FindAndReplace(wordApp, "{id:method}", keyValuePairs["{id:method}"]);
            FindAndReplace(wordApp, "{id:duration}", keyValuePairs["{id:duration}"]);
            FindAndReplace(wordApp, "{id:place}", keyValuePairs["{id:place}"]);

            string methodical = (string)keyValuePairs["{id:methodical}"];
            int k = 0;
            for (int i = 0; i < methodical.Length; i += 30)
            {
                if (i > 0)
                {
                    try
                    {
                        FindAndReplace(wordApp, "{id:methodical}", methodical.Substring(i + 1, 30) + "{id:methodical}");
                    }
                    catch (Exception e)
                    {
                        break;
                    }
                }
                else
                {
                    FindAndReplace(wordApp, "{id:methodical}", methodical.Substring(i, 30) + "{id:methodical}");
                }
                k = i;
            }
            k += 30;
            FindAndReplace(wordApp, "{id:methodical}", methodical.Substring(k + 1));
            string literature = (string)keyValuePairs["{id:literature}"];
            int l = 0;
            for (int i = 0; i < literature.Length; i += 30)
            {
                if (i > 0)
                {
                    try
                    {
                        FindAndReplace(wordApp, "{id:literature}", literature.Substring(i + 1, 30) + "{id:literature}");
                    }
                    catch (Exception e)
                    {
                        break;
                    }

                }
                else
                {
                    FindAndReplace(wordApp, "{id:literature}", literature.Substring(i, 30) + "{id:literature}");
                }
                l = i;
            }
            l += 30;
            FindAndReplace(wordApp, "{id:literature}", literature.Substring(l + 1));
            FindAndReplace(wordApp, "{id:technicalMeans}", keyValuePairs["{id:technicalMeans}"]);
            FindAndReplace(wordApp, "{id:intro}", keyValuePairs["{id:intro}"].ToString().Split(separator)[0]);
            FindAndReplace(wordApp, "{id:material}", keyValuePairs["{id:material}"]);
            FindAndReplace(wordApp, "{id:educationalQuestions}", keyValuePairs["{id:educationalQuestions}"]);
            FindAndReplace(wordApp, "{id:conclution}", keyValuePairs["{id:conclution}"].ToString().Split(separator)[0]);
            foreach (Word.Table table in doc.Tables)
            {
                Word.Range rangeTable = table.Range;
                foreach (Word.Row row in rangeTable.Rows)
                {
                    foreach (Word.Cell cell in row.Cells)
                    {
                        Word.Range range = cell.Range;
                        if (range.Text.Contains("{id:questions}"))
                        {
                            FindAndReplace(wordApp, "{id:questions}", "");
                            Word.Row temporary = row;
                            int count = 1;
                            foreach (KeyValuePair<string, string> question in (Dictionary<string, string>)keyValuePairs["{id:questions}"])
                            {
                                Object oMissing = System.Reflection.Missing.Value;
                                Word.Row newRow = table.Rows.Add(ref oMissing);
                                newRow.Cells[1].Range.Text = "2." + count;
                                newRow.Cells[2].Range.Text = question.Key;
                                newRow.Cells[3].Range.Text = question.Value.Split(separator)[0];
                                Regex regex = new Regex("^ ?[1-9].*$");
                                string questionFull = "";
                                if (regex.IsMatch(question.Key))
                                {
                                    questionFull = "Учебный вопрос. " + question.Key + " " + question.Value.Split(separator)[0] + "\r\n";
                                }
                                else
                                {
                                    questionFull = "Учебный вопрос " + count + ". " + question.Key + " " + question.Value.Split(separator)[0] + "\r\n";
                                }
                                int r = 0;
                                for (int e = 0; e < questionFull.Length; e += 30)
                                {
                                    if (e > 0)
                                    {
                                        try
                                        {
                                            FindAndReplace(wordApp, "{id:questionOfLesson}", questionFull.Substring(e, 30) + "{id:questionOfLesson}");
                                        }
                                        catch (Exception q)
                                        {
                                            break;
                                        }

                                    }
                                    else
                                    {
                                        FindAndReplace(wordApp, "{id:questionOfLesson}", questionFull.Substring(e, 30) + "{id:questionOfLesson}");
                                    }
                                    r = e;
                                }
                                r += 30;
                                FindAndReplace(wordApp, "{id:questionOfLesson}", questionFull.Substring(r) + "{id:questionOfLesson}");
                                temporary = newRow;
                                count++;
                            }
                            Object missing = System.Reflection.Missing.Value;
                            Word.Row newRowENd = table.Rows.Add(ref missing);
                            newRowENd.Cells[1].Range.Text = "3";
                            newRowENd.Cells[2].Range.Text = "Заключение";
                            newRowENd.Cells[3].Range.Text = (string)keyValuePairs["{id:conclution}"].ToString().Split(separator)[0] + " мин";
                        }
                        else if (range.Text.Contains("{id:goal}"))
                        {
                            FindAndReplace(wordApp, "{id:goal}", "");
                            int count = 1;
                            foreach (string goal in (List<string>)keyValuePairs["{id:goal}"])
                            {
                                Object oMissing = System.Reflection.Missing.Value;
                                Word.Row newRow = table.Rows.Add(ref oMissing);
                                newRow.Cells[1].Range.Text = count + "";
                                newRow.Cells[2].Range.Text = goal;
                                count++;
                            }
                        }
                    }
                }
            }
            FindAndReplace(wordApp, "{id:questionOfLesson}", "");
            WordAPI.SaveFile(doc);
            WordAPI.Close(doc);
            }
            catch (Exception e)
            {
                doc.Close();
                new ExceptionWindow()
                    .Show();
            }
            
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


