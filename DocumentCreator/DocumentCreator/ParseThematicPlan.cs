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
    class ParseThematicPlan
    {
        private static string path = Path.GetFullPath(Path.Combine(System.Reflection.Assembly.GetExecutingAssembly().Location, @"../../../../../Resources/"));
        //private static string fullPath = path + "plane.doc";
        //test path to file
        private static string fullPath = "C://plane.doc";
        private static Word.Document doc = FilesAPI.WordAPI.GetDocument(fullPath);
        private static Word.Table table = doc.Tables[2];

        //find all topics from begin index to end index (it's scope of discipline) and return map where key is name of topic and value is string with begin and end index of topic(separated with ,)
        private static Dictionary<string,string> FindByRegex(Regex regex, int beginIndex, int endIndex)
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
            resultMap[lastDiscipline] = lastValue + "," + (cells.Count - 1);
            return resultMap;
        }

        //find all discipline in file and return map where key is name of discipline and value is string with begin and end index of discipline(separated with ,)
        private static Dictionary<string, string> FindByRegexDisciplin(params Regex[] regexs)
        {
            Dictionary<string,string> resultMap = new Dictionary<string, string>();

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
                            nextDiscipline= updateRange.Text.Substring(0, updateRange.Text.Length - 4) + " " + cells[i + 1].Range.Text.Substring(0, cells[i + 1].Range.Text.Length - 2);
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
                catch(Exception e)
                {

                }
         
            }
            string lastValue=null;
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

        //Получить названия тем
        public static List<string> GetThemesOfTable()
        {
            Dictionary<string, string> resulterMap = FindByRegexDisciplin(new Regex(@"^ОВП*"), new Regex(@"^ОГП*"));
            foreach(KeyValuePair<string, string> keyValue in resulterMap)
            {
                Directory.CreateDirectory("C://output//"+keyValue.Key);
                Dictionary<string, string> resulterMapTopic = FindByRegex(new Regex(@"Тема*"), Int32.Parse(keyValue.Value.Substring(0, keyValue.Value.IndexOf(','))), Int32.Parse(keyValue.Value.Substring(keyValue.Value.IndexOf(',') + 1)));
                foreach(KeyValuePair<string, string> keyValueTopic in resulterMapTopic)
                {
                    if (keyValueTopic.Key.Length < 100)
                    {
                        string topicName = keyValueTopic.Key.Substring(0, keyValueTopic.Key.Length - 4);
                        char[] unacceptableChars = { '\\', '/', ':', '*', '?', '\"', '<', '>', '|' };
                        if (topicName.IndexOfAny(unacceptableChars)>0)
                        {
                            topicName = topicName.Substring(0,topicName.IndexOfAny(unacceptableChars));
                        }
                        Directory.CreateDirectory("C://output//" + keyValue.Key+"//"+topicName);
                    }
                    else
                    {
                        Directory.CreateDirectory("C://output//" + keyValue.Key + "//" + keyValueTopic.Key.Substring(0, 96));
                    }
                    CreateDocFileWithContenAndSave("C://output//" + keyValue.Key + "//" + keyValueTopic.Key, keyValueTopic);
                }
            }
            
 //           foreach (string theme in resulter)
 //               Directory.CreateDirectory("C://output//" + theme.Substring(0,));
 //               Console.WriteLine(theme);
            return new List<string>();
        }

        private static void CreateDocFileWithContenAndSave(string pathToDirectory, KeyValuePair<string,string> topic)
        {

        }
    }
}