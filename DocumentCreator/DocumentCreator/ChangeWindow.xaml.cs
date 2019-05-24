using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
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
        private string documentPath="";

        public ChangeWindow()
        {
            InitializeComponent();
        }
        public void initValues(Discipline discipline, Topic topic, Lesson lesson, string documentPath)
        {
            Dictionary<string, List<string>> requirementsForStudent = discipline.RequirementsForStudent;
            this.documentPath = documentPath;
            nameDiscipline.Text = discipline.Name;
            numberTopic.Text = topic.Name.Substring(0, topic.Name.IndexOf("«"));
            string temp = topic.Name.Substring(topic.Name.IndexOf("«")+1);
            topicName.Text = temp.Substring(0,temp.Length);
            numberLesson.Text = lesson.LessonInMaterialSupp;
            lessonName.Text = lesson.ThemeOfLesson;
            
            foreach(String goal in requirementsForStudent["Знать:"])
            {
                string tempGoal = goal.Replace("\v", " ");
                tempGoal = tempGoal.Replace("\r", " ");
                tempGoal = tempGoal.Replace("\a", " ");
                tempGoal = Char.ToUpper(tempGoal[0]) + tempGoal.Substring(1);
                tempGoal = tempGoal.Trim();
                selectGoal_1.Items.Add(tempGoal);
            }
            foreach (String goal in requirementsForStudent["Уметь:"])
            {
                string tempGoal = goal.Replace("\v", " ");
                tempGoal = tempGoal.Replace("\r", " ");
                tempGoal = tempGoal.Replace("\a", " ");
                tempGoal = Char.ToUpper(tempGoal[0]) + tempGoal.Substring(1);
                tempGoal = tempGoal.Trim();
                selectGoal_2.Items.Add(tempGoal);
            }
            foreach (String goal in requirementsForStudent["Владеть:"])
            {
                string tempGoal = goal.Replace("\v", " ");
                tempGoal = tempGoal.Replace("\r", " ");
                tempGoal = tempGoal.Replace("\a", " ");
                tempGoal = Char.ToUpper(tempGoal[0]) + tempGoal.Substring(1);
                tempGoal = tempGoal.Trim();
                selectGoal_3.Items.Add(tempGoal);
            }
            kind.Text = lesson.Type.Substring(0,lesson.Type.LastIndexOf(' '));
            place.Items.Add("Плац");
            place.Items.Add("Учеюный кабинет");
            place.Items.Add("Тренировочный кабинет");
            hours.Text = lesson.Minutes+" минут";
            method.Items.Add("Рассказ");
            method.Items.Add("Показ");
            method.Items.Add("Тренировка");
            if (lesson.Type.StartsWith("Лекция")) {
                introLabel.Content = "Вступительная часть:";
            }

            materialSupport.Text = lesson.MaterialSupport;
            literature.Text = lesson.Literature.Replace("\r", "; ");
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
            if (lesson.Type.StartsWith("Лекци"))
            {
                methodical.Text = discipline.MethodicalInstructionsForLecture;
            }
            else
            {
                methodical.Text = discipline.MethodicalInstructionsForRest;
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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (literature.Text.Length==0||place.Text.Length==0||materialSupport.Text.Length==0||intro_text.Text.Length==0||question1_text.Text.Length==0||conclusion_text.Text.Length==0)
            {
                ErrorWindow error = new ErrorWindow();
                error.Show();
            }
            else {
                Dictionary<string, Object> keyValuePairs = new Dictionary<string, object>();
                List<string> goals = goals_1.Text.Replace("; ", ";").Split(';').ToList<string>();
                goals.AddRange(goals_2.Text.Replace("; ", ";").Split(';').ToList<string>());
                goals.AddRange(goals_3.Text.Replace("; ", ";").Split(';').ToList<string>());
                goals.RemoveAll(x => x.Length.Equals(0));
                //List<string> goals = new List<string>();
                //goals.Add("Требоване 1");
                //goals.Add("Требоване 2");
                //goals.Add("Требоване 3");
                //goals.Add("Требоване 4");
                Dictionary<string, string> questions = new Dictionary<string, string>();
                int sumOfMinInQuestionsOfLesson = 0;
                questions.Add(questionName1.Text.Substring(0), question1_text.Text + " мин");
                sumOfMinInQuestionsOfLesson += Int32.Parse(question1_text.Text);
                if (question2_text.IsEnabled)
                {
                    questions.Add(questionName2.Text.Substring(0), question2_text.Text + " мин");
                    sumOfMinInQuestionsOfLesson += Int32.Parse(question2_text.Text);
                }
                if (question3_text.IsEnabled)
                {
                    questions.Add(questionName3.Text.Substring(0), question3_text.Text + " мин");
                    sumOfMinInQuestionsOfLesson += Int32.Parse(question3_text.Text);
                }
                if (question4_text.IsEnabled)
                {
                    questions.Add(questionName4.Text.Substring(0), question4_text.Text + " мин");
                    sumOfMinInQuestionsOfLesson += Int32.Parse(question4_text.Text);
                }
                if (question5_text.IsEnabled)
                {
                    questions.Add(questionName5.Text.Substring(0), question5_text.Text + " мин");
                    sumOfMinInQuestionsOfLesson += Int32.Parse(question5_text.Text);
                }
                keyValuePairs["{id:name}"] = nameDiscipline.Text;
                keyValuePairs["{id:theme}"] = numberTopic.Text;
                keyValuePairs["{id:themeName}"] = topicName.Text;
                keyValuePairs["{id:lesson}"] = numberLesson.Text;
                keyValuePairs["{id:lessonName}"] = lessonName.Text;
                keyValuePairs["{id:goal}"] = goals;
                keyValuePairs["{id:kind}"] = kind.Text;
                keyValuePairs["{id:method}"] = method.Text;
                keyValuePairs["{id:duration}"] = hours.Text;
                keyValuePairs["{id:place}"] = place.Text;
                keyValuePairs["{id:literature}"] = literature.Text;
                keyValuePairs["{id:intro}"] = intro_text.Text;
                keyValuePairs["{id:educationalQuestions}"] = sumOfMinInQuestionsOfLesson;
                keyValuePairs["{id:questions}"] = questions;
                keyValuePairs["{id:conclution}"] = conclusion_text.Text;
                keyValuePairs["{id:material}"] = materialSupport.Text;
                keyValuePairs["{id:methodical}"] = methodical.Text;
                keyValuePairs["{id:technicalMeans}"] = materialSupport.Text;
                UpdateDoc update = new UpdateDoc(documentPath);
                update.updateDoc(keyValuePairs);
                Close();
            }
        }

        private void Place_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            selected_Place.Text += place.SelectedItem + "; ";
        }

        private void Method_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            selected_method.Text += method.SelectedItem + "; ";
        }

        private void intro_time_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8)
                e.Handled = true;
        }

        private void question1_time_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8)
                e.Handled = true;
        }

        private void question2_time_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8)
                e.Handled = true;
        }

        private void question3_time_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8)
                e.Handled = true;
        }

        private void question4_time_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8)
                e.Handled = true;
        }

        private void question5_time_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8)
                e.Handled = true;
        }

        private void conclusion_time_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8)
                e.Handled = true;
        }

        private void Intro_combo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            intro_text.Text += intro_combo.SelectedItem + "; ";
        }

        private void Conclusion_combo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            conclusion_text.Text += conclusion_combo.SelectedItem + "; ";
        }
    }
}
