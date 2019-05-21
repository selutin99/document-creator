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
        private List<Discipline> disciplenes;

        public ParseWorkPrograming(string inputFilePath, List<Discipline> disciplenes)
        {
            this.doc = FilesAPI.WordAPI.GetDocument(inputFilePath);
            this.table = doc.Tables;
            this.disciplenes = disciplenes;
        }

        public List<Discipline> ParsePlan()
        {
            for (int j = 1; j < table.Count; j++)
            {
                Word.Range range = table[j].Range;
                Word.Cells cells = range.Cells;
                string key = "Знать";
                Dictionary<string, List<string>> requirementsForStudent = new Dictionary<string, List<string>>();
                requirementsForStudent.Add("Знать", new List<string>());
                requirementsForStudent.Add("Уметь", new List<string>());
                requirementsForStudent.Add("Владеть", new List<string>());
                List<Discipline> disciplinesCopy = new List<Discipline>(disciplenes);
                if (cells[1].Range.Text.StartsWith("В результате изучения"))
                {
                    foreach(Discipline disciplene in disciplenes)
                    {
                        if (cells[1].Range.Text.Contains(disciplene.Name))
                        {
                            string temporary = "";
                            for (int i = 1; i < cells.Count; i++)
                            {
                                Word.Cell cell = cells[i];
                                Word.Range updateRange = cell.Range;
                                string text = updateRange.Text;
                                text = text.Replace("\v", "");
                                text = text.Replace("\r", "");
                                text = text.Replace("\a", "");
                                if (text.ToLower().Equals("знать:"))
                                {
                                    temporary = "знать:";
                                    continue;
                                }
                                else if (text.ToLower().Equals("уметь:"))
                                {
                                    temporary = "уметь:";
                                    continue;
                                }
                                else if (text.ToLower().Equals("владеть:"))
                                {
                                    temporary = "владеть:";
                                    continue;
                                }
                              
                                if (temporary.Equals("знать:"))
                                {
                                    disciplinesCopy.Find(x => x.Name.Equals(disciplene.Name)).RequirementsForStudent["Знать:"].Add(text);
                                }
                                else if (temporary.Equals("уметь:"))
                                {
                                    disciplinesCopy.Find(x => x.Name.Equals(disciplene.Name)).RequirementsForStudent["Уметь:"].Add(text);
                                }
                                else if (temporary.Equals("владеть:"))
                                {
                                    disciplinesCopy.Find(x => x.Name.Equals(disciplene.Name)).RequirementsForStudent["Владеть:"].Add(text);
                                }
                            }
                        }
                    }
                    disciplenes=getMethodicalInstructionsForRest(disciplinesCopy);
                    WordAPI.Close(doc);
                    return disciplenes;
                }
            }
            WordAPI.Close(doc);
            return null;
        }
        private List<Discipline> getMethodicalInstructionsForRest(List<Discipline> disciplines)
        {
            string content = "";
            bool wasFounded = false;
            bool wasFoundedDiscipline = false;
            int k = 0;
            foreach (Word.Section section in doc.Sections)
            {
                Word.Range range = section.Range;
                int text1 = range.Text.IndexOf("Методические указания обучающимся студентам");
                if (range.Text.IndexOf("Методические указания обучающимся студентам") >= 0 || wasFounded)
                {
                    string restText = range.Text.Substring(range.Text.IndexOf("Методические указания обучающимся студентам"));
                    wasFounded = true;
                    content += restText.Substring(restText.IndexOf("Методические указания обучающимся студентам"), restText.IndexOf("Методические указания преподавателю") - restText.IndexOf("Методические указания обучающимся студентам"));
                    content = content.Trim();
                    content = content.Replace("\v", " ");
                    content = content.Replace("\r", " ");
                    content = content.Replace("\a", " ");
                    for (int i = 0; i < disciplines.Count - 1; i++)
                    {
                        try
                        {
                            if (content.IndexOf(disciplines[i + 1].Name) >= 0)
                            {
                                string temp = content.Substring(0, content.IndexOf(disciplines[i + 1].Name));
                                temp = temp.Substring(temp.IndexOf(disciplines[i].Name)+ disciplines[i].Name.Length+2);
                                disciplines[i].MethodicalInstructionsForRest = temp;
                                content = content.Substring(content.IndexOf(disciplines[i + 1].Name));
                            }
                            else if(i == disciplines.Count - 1)
                            {
                                disciplines[i+1].MethodicalInstructionsForRest = disciplines[i - 1].MethodicalInstructionsForRest;
                                disciplines[i].MethodicalInstructionsForRest = disciplines[i - 1].MethodicalInstructionsForRest;
                            }
                            else
                            {
                                disciplines[i].MethodicalInstructionsForRest = disciplines[i - 1].MethodicalInstructionsForRest;
                            }
                        }
                        catch (Exception e)
                        {
                            disciplines[i].MethodicalInstructionsForRest = content.Substring(0);
                        }



                    }
                }
            }
            return disciplines;
        }
    }
}
