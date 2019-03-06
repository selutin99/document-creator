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

        //Получить названия тем
        public static List<String> GetThemesOfTable()
        {
            List<String> resultList = new List<String>();

            Word.Table table = doc.Tables[2];
            table.Columns.AutoFit();

            var regex = new Regex(@"Тема *");
            var regException = new Regex(@"Тема и учебные вопросы занятия");

            Word.Range range = table.Range;
            Word.Cells cells = range.Cells;
            for (int i = 1; i <= cells.Count; i++)
            {
                Word.Cell cell = cells[i];
                Word.Range r2 = cell.Range;
                string txt = r2.Text;
                if (regex.IsMatch(txt))
                {
                    if (regException.IsMatch(txt))
                    {
                        continue;
                    }
                    Console.WriteLine(txt);
                    resultList.Add(txt);
                }
            }
            Console.WriteLine("Successfully finished");
            return resultList;
        }
    }
}