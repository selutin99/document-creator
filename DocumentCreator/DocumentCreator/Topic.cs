using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentCreator
{
    class Topic
    {
        string name;
        List<Lesson> lessons;

        public Topic(string name, List<Lesson> lessons)
        {
            this.Name = name;
            this.Lessons = lessons;
        }

        public string Name { get => name; set => name = value; }
        internal List<Lesson> Lessons { get => lessons; set => lessons = value; }
    }
}
