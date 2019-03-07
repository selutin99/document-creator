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
        private static string fullPath = path + "plane.doc";
        private static Word.Document doc = FilesAPI.WordAPI.GetDocument(fullPath);
        private static Word.Table table = doc.Tables[2];

        private static List<string> FindByRegex(Regex re)
        {
            List<string> resultList = new List<string>();

            Word.Range range = table.Range;
            Word.Cells cells = range.Cells;

            for (int i = 1; i <= cells.Count; i++)
            {
                Word.Cell cell = cells[i];
                Word.Range updateRange = cell.Range;
                if (re.IsMatch(updateRange.Text))
                {
                    resultList.Add(updateRange.Text);
                }
            }
            return resultList;
        }

        //Получить названия тем
        public static List<string> GetThemesOfTable()
        {             
            List<string> themes = FindByRegex(new Regex(@"Тема*"));
            themes.RemoveAt(0);
            
            return themes;
        }

        //Получить учебные вопросы
        public static Dictionary<string, int> GetDisciplines()
        {
            Dictionary<string, int> dictResulter = new Dictionary<string, int>();

            List<string> disciplines = FindByRegex(new Regex(@"ОВП*"));
            List<string> themes = FindByRegex(new Regex(@"Тема*"));

            Regex reg = new Regex(@"Тема №1*");
            int counter = 0;

            foreach (string theme in themes)
            {
                if (reg.IsMatch(theme))
                {
                    dictResulter.Add(theme, counter);
                    counter = 0;
                }
                else
                {
                    counter++;
                }
            }
            foreach(KeyValuePair<string, int> entry in dictResulter)
            {
                Console.WriteLine(entry.Key + "  ::  " + entry.Value);
            }
            return dictResulter;
        }
    }
}