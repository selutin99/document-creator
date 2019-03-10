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
        private Word.Table table;


        public ParseThematicPlan(string inputFilePath, string outputPath)
        {
            this.doc = FilesAPI.WordAPI.GetDocument(inputFilePath);
            this.table = doc.Tables[2];

            this.outputPath = outputPath;
        }

        private Dictionary<string, string> FindByRegex(Regex regex, int beginIndex, int endIndex)
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
                        if (resultMap.ContainsKey(lastDiscipline))
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
                            nextDiscipline = updateRange.Text.Substring(0, updateRange.Text.Length - 4) + " " + cells[i + 1].Range.Text.Substring(0, cells[i + 1].Range.Text.Length - 2);
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

        public List<string> LogicForParseWordAndSave()
        {
            Dictionary<string, string> resulterMap = FindByRegexDisciplin(new Regex(@"^ОВП*"), new Regex(@"^ОГП*"));
            foreach (KeyValuePair<string, string> keyValue in resulterMap)
            {
                Directory.CreateDirectory(this.outputPath + keyValue.Key);
                Dictionary<string, string> resulterMapTopic = FindByRegex(new Regex(@"Тема*"), Int32.Parse(keyValue.Value.Substring(0, keyValue.Value.IndexOf(','))), Int32.Parse(keyValue.Value.Substring(keyValue.Value.IndexOf(',') + 1)));
                foreach (KeyValuePair<string, string> keyValueTopic in resulterMapTopic)
                {
                    if (keyValueTopic.Key.Length < 100)
                    {
                        string topicName = keyValueTopic.Key.Substring(0, keyValueTopic.Key.Length - 4);
                        char[] unacceptableChars = { '\\', '/', ':', '*', '?', '\"', '<', '>', '|' };
                        if (topicName.IndexOfAny(unacceptableChars) > 0)
                        {
                            topicName = topicName.Substring(0, topicName.IndexOfAny(unacceptableChars));
                        }
                        Directory.CreateDirectory(this.outputPath + keyValue.Key + "//" + topicName);
                    }
                    else
                    {
                        Directory.CreateDirectory(this.outputPath + keyValue.Key + "//" + keyValueTopic.Key.Substring(0, 96));
                    }
                    CreateDocFileWithContenAndSave(this.outputPath + keyValue.Key + "//" + keyValueTopic.Key, keyValueTopic);
                }
            }
            //CLOSE FILE
            FilesAPI.WordAPI.Close(this.doc);
            return new List<string>();
        }

        private void CreateDocFileWithContenAndSave(string pathToDirectory, KeyValuePair<string, string> topic)
        {
            string kindOfLesson = "";
            string hours = "";
            string questionsOfLesson = "";
            string materialSupport = "";
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
                    //get count of hours
                    cell = cells[i + 1];
                    if(cell.Range.Text.Length>0)
                    {
                        hours = cell.Range.Text.Trim(charsToTrim);
                    }
                    //get questions of the lesson
                    cell = cells[i + 2];
                    questionsOfLesson = cell.Range.Text.Trim(charsToTrim);
                    //get material support
                    cell = cells[i + 3];
                    materialSupport= cell.Range.Text.Trim(charsToTrim);
                    //get literature
                    cell = cells[i + 4];
                    literature = cell.Range.Text.Trim(charsToTrim);
                    //get hours if first cell was empty
                    if (hours == "")
                    {
                        cell = cells[i + 5];
                        hours= cell.Range.Text.Trim(charsToTrim);
                    }
                    CreateFile(pathToDirectory, kindOfLesson, materialSupport, literature);
                    i += 5;
                }
            }
        }

        //Создаёт Word файл
        private void CreateFile(string pathToDirectory, string kind, string materialSupport, string literature)
        {
            string path = Path.GetFullPath(Path.Combine(System.Reflection.Assembly.GetExecutingAssembly().Location, @"../../../../../Resources/"));
            string fileName = path+"theme.doc";
            string outputFileName = pathToDirectory+kind+".doc";

            fileName = CleanFileName(fileName);
            outputFileName = CleanOutput(outputFileName);

            outputFileName = outputFileName.Replace("//", "\\");
            outputFileName = outputFileName.Replace("\"", "");
            outputFileName = outputFileName.Replace("\\\\", "\\");

            File.Copy(@fileName, @outputFileName);
        }

        private static string CleanFileName(string fileName)
        {
            return string.Join("", fileName.Split(Path.GetInvalidPathChars()));
        }

        private static string CleanOutput(string fileName)
        {
            return string.Join("\\", fileName.Split(Path.GetInvalidPathChars()));
        }
    }
}