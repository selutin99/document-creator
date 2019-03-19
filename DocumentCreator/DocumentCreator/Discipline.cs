﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentCreator
{
    class Discipline
    {
        string name;
        List<Topic> topics;

        public Discipline(string name, List<Topic> topics)
        {
            this.Name = name;
            this.Topics = topics;
        }

        public string Name { get => name; set => name = value; }
        internal List<Topic> Topics { get => topics; set => topics = value; }
    }
}