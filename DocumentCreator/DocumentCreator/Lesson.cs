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
        public string Type { get { return type; } set { type = value; } }
        public string Minutes { get { return minutes; } set { minutes = value; } }
        public string MaterialSupport { get { return materialSupport; } set { materialSupport = value; } }
        public string LessonInMaterialSupp { get { return lessonInMaterialSupp; } set { lessonInMaterialSupp = value; } }
        public string ThemeOfLesson { get { return themeOfLesson; } set { themeOfLesson = value; } }
        public List<string> Questions { get { return questions; } set { questions = value; } }
        public string Literature { get { return literature; } set { literature = value; } }
    }
}
