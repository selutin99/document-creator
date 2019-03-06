using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Word = Microsoft.Office.Interop.Word;

namespace DocumentCreator
{
    class ParseThematicPlan
    {
        private static string path = Path.GetFullPath(Path.Combine(System.Reflection.Assembly.GetExecutingAssembly().Location, @"../../../../Resources/"));
        private static string fullPath = path + "темплан_1.docx";
        private static Word.Document doc = FilesAPI.WordAPI.GetDocument(fullPath);

        //Получить названия тем
        public static void GetThemes()
        {
            Word.Table table = doc.Tables[2];
            for (int i = 5; i <= 8; i++)
            {
                for (int j = 1; j <= 7; j++)
                {
                    var cell = table.Cell(i, j);
                    if(String.IsNullOrEmpty(cell.Range.Text) && (j == 1 || j == 2))
                    {
                        Console.WriteLine(cell.Range.Text);
                    }
                }
            }
            Console.WriteLine("I am finish");
            //FilesAPI.WordAPI.SaveFile(doc);
            FilesAPI.WordAPI.Close(doc);
        }
    }
}
