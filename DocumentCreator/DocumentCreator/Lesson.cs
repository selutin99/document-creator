using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentCreator
{
    class Lesson
    {
        string type;
        string hours;
        string materialSupport;
        string lessonInMaterialSupp;
        string themeOfLesson;
        List<string> questions;
        string literature;
<<<<<<< HEAD
        public string Type { get => type; set => type = value; }
        public string Hours { get => hours; set => hours = value; }
        public string MaterialSupport { get => materialSupport; set => materialSupport = value; }
        public string LessonInMaterialSupp { get => lessonInMaterialSupp; set => lessonInMaterialSupp = value; }
        public string ThemeOfLesson { get => themeOfLesson; set => themeOfLesson = value; }
        public List<string> Questions { get => questions; set => questions = value; }
        public string Literature { get => literature; set => literature = value; }
=======
        public string Type { get { return type; } set { type = value; } }
        public string Hours { get { return hours; } set { hours = value; } }
        public string Content { get { return content; } set { content = value; } }
        public string MaterialSupport { get { return materialSupport; } set { materialSupport = value; } }
        public string Literature { get { return literature; } set { literature = value; } }
>>>>>>> master
    }
}
