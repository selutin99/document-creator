using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentCreator
{
    public class Lesson
    {
        string type;
        string minutes;
        string materialSupport;
        string lessonInMaterialSupp;
        string themeOfLesson;
        List<string> questions;
        string literature;
        public string Type { get => type; set => type = value; }
        public string Minutes { get => minutes; set => minutes = value; }
        public string MaterialSupport { get => materialSupport; set => materialSupport = value; }
        public string LessonInMaterialSupp { get => lessonInMaterialSupp; set => lessonInMaterialSupp = value; }
        public string ThemeOfLesson { get => themeOfLesson; set => themeOfLesson = value; }
        public List<string> Questions { get => questions; set => questions = value; }
        public string Literature { get => literature; set => literature = value; }
    }
}
