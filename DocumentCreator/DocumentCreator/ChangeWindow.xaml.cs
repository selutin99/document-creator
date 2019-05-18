using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace DocumentCreator
{
    /// <summary>
    /// Логика взаимодействия для ChangeWindow.xaml
    /// </summary>
    public partial class ChangeWindow : Window
    {

        public ChangeWindow()
        {
            InitializeComponent();
        }
        public void initValues(Discipline discipline, Topic topic, Lesson lesson, Dictionary<string, List<string>> requirementsForStudent)
        {
            nameDiscipline.Text = discipline.Name;
            topicName.Text = topic.Name;
            lessonName.Text = lesson.ThemeOfLesson;
            foreach(String goal in requirementsForStudent["Знать"])
            {
                selectGoal_1.Items.Add(goal);
            }
            foreach (String goal in requirementsForStudent["Уметь"])
            {
                selectGoal_2.Items.Add(goal);
            }
            foreach (String goal in requirementsForStudent["Владеть"])
            {
                selectGoal_3.Items.Add(goal);
            }
            kind.Text = lesson.Type;
            place.Items.Add("Плац");
            place.Items.Add("Учеюный кабинет");
            place.Items.Add("Тренировочный кабинет");
            hours.Text = lesson.Hours;
            materialSupport.Text = lesson.LessonInMaterialSupp;
            literature.Text = lesson.Literature;
            for(int i=0;i< lesson.Questions.Count; i++)
            {
                if (i == 0)
                {
                    questionName1.Text = lesson.Questions[i];
                    question1_text.IsEnabled = true;

                }
                else if(i ==1)
                {
                    questionName2.Text = lesson.Questions[i];
                    question2_text.IsEnabled = true;
                }
                else if (i == 2)
                {
                    questionName3.Text = lesson.Questions[i];
                    question3_text.IsEnabled = true;
                }
                else if (i == 3)
                {
                    questionName4.Text = lesson.Questions[i];
                    question4_text.IsEnabled = true;
                }
                else if (i == 4)
                {
                    questionName5.Text = lesson.Questions[i];
                    question5_text.IsEnabled = true;
                }
            }



        }

        //private void ComboBox_MouseDown(object sender, MouseButtonEventArgs e)
        //{
           
        //}

        //private void SelectGoal_2_MouseDown(object sender, MouseButtonEventArgs e)
        //{
            
        //}

        //private void SelectGoal_3_MouseDown(object sender, MouseButtonEventArgs e)
        //{
            
        //}

        private void SelectGoal_3_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            goals_3.Text += selectGoal_3.SelectedItem + "; ";
        }

        private void SelectGoal_2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            goals_2.Text += selectGoal_2.SelectedItem + "; ";
        }

        private void SelectGoal_1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            goals_1.Text += selectGoal_1.SelectedItem + "; ";
        }

        //private void Place_SelectionChanged(object sender, SelectionChangedEventArgs e)
        //{

        //}
    }
}
