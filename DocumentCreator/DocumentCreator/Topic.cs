using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentCreator
{
    public class Topic
    {
        string name;
        string numberTopic;
        string cutName;
        List<Lesson> lessons;

        public Topic(string name, List<Lesson> lessons)
        {
            this.Name = name;
            this.Lessons = lessons;
        }

        public string Name { get { return name; } set { name = value; } }
        internal List<Lesson> Lessons { get { return lessons; } set { lessons = value; } }
        public string NumberTopic { get { return numberTopic; } set { numberTopic = value; } }
        public string CutName { get { return cutName; } set { cutName = value; } }
    }
}
