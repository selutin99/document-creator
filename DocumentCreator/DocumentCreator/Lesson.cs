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
        string content;
        string materialSupport;
        string literature;
        public string Type { get => type; set => type = value; }
        public string Hours { get => hours; set => hours = value; }
        public string Content { get => content; set => content = value; }
        public string MaterialSupport { get => materialSupport; set => materialSupport = value; }
        public string Literature { get => literature; set => literature = value; }
    }
}
