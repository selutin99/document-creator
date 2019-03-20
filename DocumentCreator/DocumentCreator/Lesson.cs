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
        public string Type { get { return type; } set { type = value; } }
        public string Hours { get { return hours; } set { hours = value; } }
        public string Content { get { return content; } set { content = value; } }
        public string MaterialSupport { get { return materialSupport; } set { materialSupport = value; } }
        public string Literature { get { return literature; } set { literature = value; } }
    }
}
