using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace DocumentCreator
{
    internal class ParseThematicPlan
    {
        private string outputPath;
        private Word.Document doc;
        private Word.Table table = null;

        public string GetOutputPath()
        {
            return this.outputPath;
        }

        public ParseThematicPlan(string inputFilePath, string outputPath)
        {
            this.doc = FilesAPI.WordAPI.GetDocument(inputFilePath);
            foreach(Word.Table tbl in doc.Tables)
            {
                int k = 0;
                if (this.table == null)
                {
                    foreach (Word.Cell cell in tbl.Range.Cells)
                    {
                        if (k >= 200)
                        {
                            break;
                        }

                        if (cell.Range.Text.Contains("Тема и учебные вопросы занятия"))
                        {
                            this.table = tbl;
                        }
                        k++;
                    }
                }
                else
                {
                    break;
                }
                
                
            }

            this.outputPath = outputPath;
        }

        private Dictionary<string, string> FindByRegexTopics(Regex regex, int beginIndex, int endIndex)
        {
            Dictionary<string, string> resultMap = new Dictionary<string, string>();

            Word.Range range = table.Range;
            Word.Cells cells = range.Cells;
            string lastDiscipline = null;
            string nextDiscipline = null;
            for (int i = beginIndex; i <= endIndex; i++)
            {
                Word.Cell cell = cells[i];
                Word.Range updateRange = cell.Range;
                try
                {
                    if (regex.IsMatch(updateRange.Text))
                    {
                        nextDiscipline = updateRange.Text;
                        if (lastDiscipline == null)
                        {
                            lastDiscipline = nextDiscipline;
                        }
                        if (resultMap.ContainsKey(lastDiscipline)&&resultMap[lastDiscipline].IndexOf(',')>0)
                        {
                            string value;
                            resultMap.TryGetValue(lastDiscipline, out value);
                            value = value.Substring(0, value.IndexOf(','));
                            resultMap[lastDiscipline] = value + "," + i;
                            resultMap.Add(nextDiscipline, i.ToString());

                        }
                        else if (resultMap.ContainsKey(lastDiscipline))
                        {
                            string value;
                            resultMap.TryGetValue(lastDiscipline, out value);
                            resultMap[lastDiscipline] = value + "," + i;
                            resultMap.Add(nextDiscipline, i.ToString());
                        }
                        else
                        {
                            resultMap.Add(nextDiscipline, i.ToString());
                        }
                    }
                    lastDiscipline = nextDiscipline;

                }
                catch (Exception e)
                {

                }

            }
            string lastValue;
            resultMap.TryGetValue(lastDiscipline, out lastValue);
            resultMap[lastDiscipline] = lastValue + "," + endIndex;
            return resultMap;
        }

        //find all discipline in file and return map where key is name of discipline and value is string with begin and end index of discipline(separated with ,)
        private Dictionary<string, string> FindByRegexDisciplin(params Regex[] regexs)
        {
            Dictionary<string, string> resultMap = new Dictionary<string, string>();

            Word.Range range = table.Range;
            Word.Cells cells = range.Cells;
            string lastDiscipline = null;
            string nextDiscipline = null;
            for (int i = 1; i <= cells.Count; i++)
            {
                Word.Cell cell = cells[i];
                Word.Range updateRange = cell.Range;
                try
                {
                    foreach (Regex re in regexs)
                    {
                        if (re.IsMatch(updateRange.Text))
                        {
                            nextDiscipline = updateRange.Text.Substring(0, updateRange.Text.Length);
                            nextDiscipline = nextDiscipline.Trim();
                            nextDiscipline = nextDiscipline.Replace("\v", "");
                            nextDiscipline = nextDiscipline.Replace("\r", "");
                            nextDiscipline = nextDiscipline.Replace("\a", "");
                            nextDiscipline = nextDiscipline.Replace("  ", " ");
                            if (lastDiscipline == null)
                            {
                                lastDiscipline = nextDiscipline;
                            }
                            if (resultMap.ContainsKey(lastDiscipline))
                            {
                                string value;
                                resultMap.TryGetValue(lastDiscipline, out value);
                                if (value.IndexOf(',') > 0)
                                {
                                    value = value.Substring(0, value.IndexOf(','));
                                    resultMap[lastDiscipline] = value + "," + i;
                                    resultMap.Add(nextDiscipline, i.ToString());
                                }
                                else
                                {
                                    resultMap[lastDiscipline] = value + "," + i;
                                    resultMap.Add(nextDiscipline, i.ToString());
                                }
                            }
                            else
                            {
                                resultMap.Add(nextDiscipline, i.ToString());
                            }
                        }
                        lastDiscipline = nextDiscipline;
                    }
                }
                catch (Exception e)
                {
                }
            }
            string lastValue = null;
            if (lastDiscipline == null)
            {
                foreach(Word.Paragraph paragraph in doc.Paragraphs)
                {
                    Regex regex1 = new Regex(@"^*ОВП*");
                    Regex regex2 = new Regex(@"^*ОГП*");
                    Regex regex3 = new Regex(@"^*ВТП*");
                    string text = paragraph.Range.Text;
                    if(regex1.IsMatch(text))
                    {
                        text = text.Substring(text.IndexOf("OВП")).Trim();
                        resultMap.Add(text, "9," + (cells.Count - 1));
                        return resultMap;
                    }
                    if (regex2.IsMatch(text))
                    {
                        text = text.Substring(text.IndexOf("OГП")).Trim();
                        resultMap.Add(text, "9," + (cells.Count - 1));
                        return resultMap;
                    }
                    if (regex3.IsMatch(text))
                    {
                        text = text.Substring(text.IndexOf("ВТП")).Trim();
                        resultMap.Add(text, "9," + (cells.Count - 1));
                        return resultMap;
                    }

                }
            }
            resultMap.TryGetValue(lastDiscipline, out lastValue);
            if (lastValue.IndexOf(',') > 0)
            {
                lastValue = lastValue.Substring(0, lastValue.IndexOf(','));
                resultMap[lastDiscipline] = lastValue + "," + (cells.Count - 1);
            }
            else
            {
                resultMap[lastDiscipline] = lastValue + "," + (cells.Count - 1);
            }

            return resultMap;
        }

        private List<Discipline> GetAllDisciplinesWithContent()
        {
            Dictionary<string, string> resulterMap = FindByRegexDisciplin(new Regex(@"^ОВП*"), new Regex(@"^ОГП*"),new Regex(@"^ВТП*"));
            List<Discipline> disciplines = new List<Discipline>();
            
            foreach (KeyValuePair<string, string> keyValue in resulterMap)
            {
                Discipline discipline = new Discipline(keyValue.Key, new List<Topic>());

                
                
                Dictionary<string, string> resulterMapTopic = FindByRegexTopics(new Regex(@"Тема*"), Int32.Parse(keyValue.Value.Substring(0, keyValue.Value.IndexOf(','))), Int32.Parse(keyValue.Value.Substring(keyValue.Value.IndexOf(',') + 1)));
                foreach (KeyValuePair<string, string> keyValueTopic in resulterMapTopic)
                {
                    string topicName;
                    if (keyValueTopic.Key.Length < 100)
                    {
                        topicName = keyValueTopic.Key.Substring(0, keyValueTopic.Key.Length);
                        topicName = topicName.Trim();
                        topicName = topicName.Replace("\r", "");
                        topicName = topicName.Replace("\a", "");
                        char[] unacceptableChars = { '\\', '/', ':', '*', '?', '\"', '<', '>', '|' };
                        if (topicName.IndexOfAny(unacceptableChars) > 0)
                        {
                            topicName = topicName.Substring(0, topicName.IndexOfAny(unacceptableChars));
                        }
                    }
                    else
                    {
                        topicName = keyValueTopic.Key.Substring(0, 96);
                        topicName = topicName.Trim();
                        topicName = topicName.Replace("\r", "");
                        topicName = topicName.Replace("\a", "");
                    }
                    discipline.Topics.Add(new Topic(topicName, GetLessonsByTopic(keyValueTopic)));
                }
                disciplines.Add(discipline);
            }
            disciplines=getMethodicalInstructionsForLecture(disciplines);
            disciplines = replaceLiterature(disciplines);
            //CLOSE FILE
            FilesAPI.WordAPI.Close(this.doc);
            //парсим метод указания для лекций
            
            return disciplines;
        }

        private List<Lesson> GetLessonsByTopic(KeyValuePair<string, string> topic)
        {
            List<Lesson> lessons = new List<Lesson>();
            string kindOfLesson = "";
            string minutes = "";
            string questionsOfLesson = "";
            string materialSupport = "";
            string lessonInMaterialSupp = "";
            string themeOfLesson = "";
            List<string> questions = new List<string>();
            string literature = "";
            Word.Range range = table.Range;
            Word.Cells cells = range.Cells;
            Regex regex = new Regex(@"^Лекция|^Самостоя|^Группов|^Практичес|^Трениров");
            char[] charsToTrim = { '\a', '\r' };
            for (int i=Int32.Parse(topic.Value.Substring(0,topic.Value.IndexOf(',')))+1;i< Int32.Parse(topic.Value.Substring(topic.Value.IndexOf(',') + 1)); i++)
            {
                Word.Cell cell = cells[i];
                Word.Range updateRange = cell.Range;
                string text = updateRange.Text;
                if (regex.IsMatch(text))
                {
                    kindOfLesson = text.Trim(charsToTrim);
                    kindOfLesson = kindOfLesson.Replace("\r","");
                    //get count of hours
                    cell = cells[i + 1];
                    if(cell.Range.Text.Length>0)
                    {
                        minutes = cell.Range.Text.Trim(charsToTrim);
                    }
                    //get questions of the lesson
                    cell = cells[i + 2];
                    questionsOfLesson = cell.Range.Text.Trim(charsToTrim);
                    int a = questionsOfLesson.IndexOf("«");
                    if (questionsOfLesson.IndexOf("«") > 20)
                    {
                        lessonInMaterialSupp = "Ошибка в темплане";
                        themeOfLesson = "Ошибка в темплане";
                        questions = getQuestions(questionsOfLesson);
                    }
                    else
                    {
                        lessonInMaterialSupp = questionsOfLesson.Substring(0, questionsOfLesson.IndexOf("«"));
                        try
                        {
                            themeOfLesson = questionsOfLesson.Substring(questionsOfLesson.IndexOf("«") + 1, questionsOfLesson.IndexOf("»") - questionsOfLesson.IndexOf("«")-1);
                        }
                        catch (Exception e)
                        {
                            themeOfLesson = questionsOfLesson.Substring(questionsOfLesson.IndexOf("«") + 1, questionsOfLesson.IndexOf(".") - questionsOfLesson.IndexOf("«")-1);
                        }
                        questions = getQuestions(questionsOfLesson.Substring(questionsOfLesson.IndexOf("«")));
                    }
                    
                    //get material support
                    cell = cells[i + 3];
                    materialSupport= cell.Range.Text.Trim(charsToTrim);
                   
                    //get literature
                    cell = cells[i + 4];
                    literature = cell.Range.Text.Trim(charsToTrim);
                    
                    //get hours if first cell was empty
                    if (minutes == "")
                    {
                        cell = cells[i + 5];
                        minutes= cell.Range.Text.Trim(charsToTrim);
                    }
                    Lesson lesson = new Lesson();
                    lesson.Type = kindOfLesson;
                    lesson.Literature = literature;
                    lesson.LessonInMaterialSupp = lessonInMaterialSupp;
                    lesson.ThemeOfLesson = themeOfLesson;
                    lesson.Questions = questions;
                    lesson.MaterialSupport = materialSupport;
                    //ПРО САМОСТОЯТЕЛЬНУ РАБОТУ СПРОСИТЬ СКОЛЬКО ТАМ МИНУТ БУДЕТ
                    if(lesson.Type.Contains("Лекци")|| lesson.Type.Contains("Групповое")|| lesson.Type.Contains("Практичес"))
                    {
                        int count = Int32.Parse(minutes) * 45;
                        minutes = count.ToString();
                    }
                    else if (lesson.Type.Contains("Трениров"))
                    {
                        minutes = "30";
                    }
                    lesson.Minutes = minutes;
                    lessons.Add(lesson);
                    i += 5;
                }
            }
            return lessons;
        }

        public List<Discipline> ParseThematicPlanAndCreateDirectories()
        {
            List<Discipline> disciplines = GetAllDisciplinesWithContent();
            return disciplines;
        }
         
        private List<string> getQuestions(string questions)
        {
            string temporary;
            string question;
            List<string> listQuestions = new List<string>();
            try
            {
                temporary = questions.Substring(questions.IndexOf("1."));
                if (temporary.IndexOf("2.") > 0)
                {
                    question = temporary.Substring(0, temporary.IndexOf("2."));
                    question = question.Trim();
                    listQuestions.Add(question);
                    temporary = temporary.Substring(temporary.IndexOf("2."));
                }
                if (temporary.IndexOf("3.") > 0)
                {
                    question = temporary.Substring(0, temporary.IndexOf("3."));
                    question = question.Trim();
                    listQuestions.Add(question);
                    temporary = temporary.Substring(temporary.IndexOf("3."));
                }
                if (temporary.IndexOf("4.") > 0)
                {
                    question = temporary.Substring(0, temporary.IndexOf("4."));
                    question = question.Trim();
                    listQuestions.Add(question);
                    temporary = temporary.Substring(temporary.IndexOf("4."));
                }
                if (temporary.IndexOf("5.") > 0)
                {
                    question = temporary.Substring(0, temporary.IndexOf("5."));
                    question = question.Trim();
                    listQuestions.Add(question);
                    temporary = temporary.Substring(temporary.IndexOf("5."));
                }
                //    while (temporary.IndexOf("\r") > 0) { 
                //    if (temporary.LastIndexOf("\r") > 0)
                //    {
                //        question = temporary.Substring(0, temporary.IndexOf("\r") + 1);
                //        question = question.Trim();
                //        listQuestions.Add(question);
                //        temporary = temporary.Substring(temporary.IndexOf("\r") + 1);
                //    }
                //    else
                //    {
                //        listQuestions.Add(temporary);
                //        return listQuestions;
                //    }
                //}


                listQuestions.Add(temporary);


                return listQuestions;
            }
            catch(Exception e)
            {
                temporary = questions.Substring(questions.IndexOf("\r")+1);
                while (temporary.IndexOf("\r") > 0)
                {
                    if (temporary.LastIndexOf("\r") > 0)
                    {
                        question = temporary.Substring(0, temporary.IndexOf("\r") + 1);
                        question = question.Trim();
                        listQuestions.Add(question);
                        temporary = temporary.Substring(temporary.IndexOf("\r") + 1);
                    }
                    else
                    {
                        listQuestions.Add(temporary);
                        return listQuestions;
                    }
                }

                listQuestions.Add(temporary);
                return listQuestions;
            }
        }
        private List<Discipline> getMethodicalInstructionsForLecture(List<Discipline> disciplines)
        {
            string content = "";
            bool wasFounded = false;
            bool wasFoundedDiscipline = false;
            int k = 0;
            foreach(Word.Section section in doc.Sections)
            {
                Word.Range range = section.Range;
                int text1 = range.Text.IndexOf("Организационно-методические указания");
                if (range.Text.IndexOf("Организационно-методические указания") >= 0 || wasFounded)
                {
                    string restText = range.Text.Substring(range.Text.IndexOf("Организационно-методические указания") + "Организационно-методические указания".Length);
                    wasFounded = true;
                    if (restText.IndexOf("Часть 1") < 0)
                    {
                        content += restText.Substring(0, restText.IndexOf("Часть 2"));
                    }
                    else
                    {
                        content += restText.Substring(restText.IndexOf("Часть 1") + "Часть 1".Length, restText.IndexOf("Часть 2") - restText.IndexOf("Часть 1"));
                    }
                    content = content.Trim();
                    content = content.Replace("\v", " ");
                    content = content.Replace("\r", " ");
                    content = content.Replace("\a", " ");
                    if (disciplines.Count == 1)
                    {
                        disciplines[0].MethodicalInstructionsForLecture = content.Substring(0);
                        return disciplines;
                    }
                    {
                        for (int i = 0; i < disciplines.Count - 1; i++)
                        {
                            try
                            {
                                if (content.IndexOf(disciplines[i + 1].Name) >= 0)
                                {
                                    string temp = content.Substring(0, content.IndexOf(disciplines[i + 1].Name));
                                    temp = temp.Substring(temp.IndexOf(disciplines[i].Name) + disciplines[i].Name.Length + 2);
                                    disciplines[i].MethodicalInstructionsForLecture = temp;
                                    content = content.Substring(content.IndexOf(disciplines[i + 1].Name));
                                }
                                else
                                {
                                    disciplines[i].MethodicalInstructionsForLecture = disciplines[i - 1].MethodicalInstructionsForLecture;
                                }
                            }
                            catch (Exception e)
                            {
                                disciplines[i].MethodicalInstructionsForLecture = content.Substring(0);
                            }



                        }
                    }
                }
            }
            return disciplines;
        }
        private List<Discipline> replaceLiterature(List<Discipline> disciplines)
        {
            Word.Table tableWithLiterature = null;
            foreach (Word.Table tbl in doc.Tables)
            {
                int k = 0;
                if (tableWithLiterature == null)
                {
                    foreach (Word.Cell cell in tbl.Range.Cells)
                    {
                        if (k >= 3)
                        {
                            break;
                        }

                        if (cell.Range.Text.Contains("ЛИТЕРАТУРА")|| cell.Range.Text.Contains("литература")|| cell.Range.Text.Contains("Литература"))
                        {
                            tableWithLiterature = tbl;
                        }
                        k++;
                    }
                }
                else
                {
                    break;
                }
            }
            Dictionary<string, string> mainLiterature = new Dictionary<string, string>();
            Dictionary<string, string> additionalLiterature = new Dictionary<string, string>();
            string temp = "";
            for (int i=1; i<tableWithLiterature.Range.Cells.Count;i++)
            {
                if (temp.Equals("main")&& (!tableWithLiterature.Range.Cells[i].Range.Text.Contains(" Дополнительная")))
                {
                    string replaceKey = tableWithLiterature.Range.Cells[i].Range.Text;
                    replaceKey = replaceKey.Trim();
                    replaceKey = replaceKey.Replace("\v", "");
                    replaceKey = replaceKey.Replace("\r", "");
                    replaceKey = replaceKey.Replace("\a", "");
                    replaceKey = replaceKey.Replace(".", "");
                    if (replaceKey.Equals(""))
                    {
                        replaceKey = mainLiterature.Keys.Count + 1 + "";
                    }
                    string replaceValue=tableWithLiterature.Range.Cells[i + 1].Range.Text;
                    replaceValue = replaceValue.Trim();
                    replaceValue = replaceValue.Replace("\v", " ");
                    replaceValue = replaceValue.Replace("\r", " ");
                    replaceValue = replaceValue.Replace("\a", " ");
                    mainLiterature.Add(replaceKey, replaceValue);
                    i = i + 1;
                }
                else if (temp.Equals("additional"))
                {
                    string replaceKey = tableWithLiterature.Range.Cells[i].Range.Text;
                    replaceKey = replaceKey.Trim();
                    replaceKey = replaceKey.Replace("\v", "");
                    replaceKey = replaceKey.Replace("\r", "");
                    replaceKey = replaceKey.Replace("\a", "");
                    replaceKey = replaceKey.Replace(".", "");
                    if (replaceKey.Equals(""))
                    {
                        replaceKey = additionalLiterature.Keys.Count + 1 + "";
                    }
                    string replaceValue = tableWithLiterature.Range.Cells[i + 1].Range.Text;
                    replaceValue = replaceValue.Trim();
                    replaceValue = replaceValue.Replace("\v", " ");
                    replaceValue = replaceValue.Replace("\r", " ");
                    replaceValue = replaceValue.Replace("\a", " ");
                    additionalLiterature.Add(replaceKey, replaceValue);
                    i = i + 1;
                }
                if(tableWithLiterature.Range.Cells[i].Range.Text.Contains(" Основная"))
                {
                    temp = "main";
                }
                else if (tableWithLiterature.Range.Cells[i].Range.Text.Contains(" Дополнительная") || (!temp.Equals("") && tableWithLiterature.Range.Cells[i].Range.Text.Contains(" Дополнительная")))
                {
                    temp = "additional";
                }

            }
            List<Discipline> disciplinesCopy = new List<Discipline>();
            string finalStringConcat = "";
            for (int i= 0; i < disciplines.Count; i++)
            {
                for (int j = 0; j < disciplines[i].Topics.Count; j++)
                {
                    for (int k = 0; k < disciplines[i].Topics[j].Lessons.Count; k++)
                    {
                        
                        string[] mas= disciplines[i].Topics[j].Lessons[k].Literature.Split(';');
                        disciplines[i].Topics[j].Lessons[k].Literature="";
                        for (int e = 0; e < mas.Length-1; e++)
                        {
                            string literature = mas[e];
                            literature = literature.Trim();
                            if (literature.IndexOf(",") < literature.IndexOf(" ")&& literature.IndexOf(",")!=-1)
                            {
                                literature = literature.Split(',')[0];
                                finalStringConcat = mas[e].Substring(mas[e].IndexOf(","));
                            }
                            else
                            {
                                literature = literature.Split(' ')[0];
                                finalStringConcat = mas[e].Substring(mas[e].IndexOf(" "));
                            }
                            literature = literature.Replace("\v", "");
                            literature = literature.Replace("\r", "");
                            literature = literature.Replace("\a", "");
                            if (literature.StartsWith("А") || literature.StartsWith("A"))
                            {
                                literature = literature.Substring(1);
                                disciplines[i].Topics[j].Lessons[k].Literature += mainLiterature[literature] + finalStringConcat + ";";
                            }
                            else
                            {
                                literature = literature.Substring(1);
                                disciplines[i].Topics[j].Lessons[k].Literature += additionalLiterature[literature] + finalStringConcat + ";";
                            }
                        }
                    }
                }
            }
            return disciplines;
        }
    }
}