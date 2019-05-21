using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentCreator
{
    public class Discipline
    {
        string name;
        Dictionary<string, List<string>> requirementsForStudent;
        List<Topic> topics;
        string methodicalInstructionsForLecture;
        string methodicalInstructionsForRest;

        public Discipline(string name, List<Topic> topics)
        {
            this.Name = name;
            this.Topics = topics;
            requirementsForStudent = new Dictionary<string, List<string>>();
            requirementsForStudent.Add("Знать:", new List<string>());
            requirementsForStudent.Add("Уметь:", new List<string>());
            requirementsForStudent.Add("Владеть:", new List<string>());
        }

        public string Name { get { return name; } set { name = value; } }
        public string MethodicalInstructionsForLecture { get { return methodicalInstructionsForLecture; } set { methodicalInstructionsForLecture = value; } }
        public string MethodicalInstructionsForRest { get { return methodicalInstructionsForRest; } set { methodicalInstructionsForRest = value; } } 
        internal List<Topic> Topics { get { return topics; } set { topics = value; } }
        public Dictionary<string, List<string>> RequirementsForStudent { get { return requirementsForStudent; } set { requirementsForStudent = value; } }
    }
}
