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
        private Word.Document doc = null;
        private Word.Table table = null;

        public ParseThematicPlan(string filePath)
        {
            this.doc = FilesAPI.WordAPI.GetDocument(filePath);
            this.table = doc.Tables[2];
        }

        private Dictionary<int, string> FindByRegexInTableDict(Regex re)
        {
            Dictionary<int, string> resultDict = new Dictionary<int, string>();

            Word.Range range = table.Range;
            Word.Cells cells = range.Cells;

            int id = 1; 

            for (int i = 1; i <= cells.Count; i++)
            {
                Word.Cell cell = cells[i];
                Word.Range updateRange = cell.Range;
                if (re.IsMatch(updateRange.Text))
                {
                    resultDict.Add(id, updateRange.Text);
                    id++;
                }
            }
            return resultDict;
        }

        private List<string> FindByRegexInTable(Regex re)
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

        //Получить список дисциплин
        public List<string> GetDisciplines()
        {
            return FindByRegexInTable(new Regex(@"ОВП.*")).
                   Concat(FindByRegexInTable(new Regex(@"ОГП Дисциплина*"))).
                   ToList();
        }

        //Получить названия тем
        public List<string> GetThemesOfTable()
        {             
            List<string> themes = FindByRegexInTable(new Regex(@"Тема*"));
            themes.RemoveAt(0);
            
            return themes;
        }

        //Получить список домашних заданий
        // ключ - № п/п
        // значение - домашнее задание
        public Dictionary<int, string> GetHomeWorks()
        {
            Dictionary<int, string> resulter = FindByRegexInTableDict(new Regex(@"А(\s|)\d{1,}"));
            return resulter;
        }

        //Получить виды учебных занятий
        // ключ - № п/п
        // значение - вид учебного занятия
        public Dictionary<int, string> GetTypeStudies()
        {
            Regex reg = new Regex(@"(Лекция|Самостоятельная|Групповое|Практическое|Практическая|Семинар|Тренировка(\s| )№)");
            Dictionary<int, string> resulter = FindByRegexInTableDict(reg);
            return resulter;
        }

        //Получить виды учебных занятий
        // ключ - № п/п
        // значение - материальное обеспечение
        public Dictionary<int, string> GetMaterialSecurity()
        {
            Regex reg = new Regex(@"(Презентация по|Компьютер|Строевой плац)");
            Dictionary<int, string> resulter = FindByRegexInTableDict(reg);
            return resulter;
        }

        //Получить темы и учебные занятия
        // ключ - № п/п
        // значение - материальное обеспечение
        public Dictionary<int, string> GetSessions()
        {
            Dictionary<int, string> resulter = FindByRegexInTableDict(new Regex(@"Занятие"));
            return resulter;
        }
    }
}