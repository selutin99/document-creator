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
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Path = System.IO.Path;

namespace DocumentCreator
{
    /// <summary>
    /// Логика взаимодействия для ChangeWindow.xaml
    /// </summary>
    public partial class ChangeWindow : Window
    {
        private string documentPath="";
        private string documentPathPlan = "";

        public ChangeWindow()
        {
            InitializeComponent();
        }
        public void initValues(Discipline discipline, Topic topic, Lesson lesson, string documentPath,string documentPathPlan)
        {
            Word.Document docWithGoals;
            Word.Table tableInDoc = null;
            string tem = Path.Combine(System.Reflection.Assembly.GetExecutingAssembly().Location, @"../../../../../Resources/ВоспитательныеЦелиИФразыДополнения.docx");
            docWithGoals = FilesAPI.WordAPI.GetDocument(Path.GetFullPath(Path.Combine(System.Reflection.Assembly.GetExecutingAssembly().Location, @"../../../../../Resources/ВоспитательныеЦелиИФразыДополнения.docx")));
            tableInDoc = docWithGoals.Tables[2];
            List<string> additionalPhraze = new List<string>();
            List<string> placeOfEducations = new List<string>();
            List<string> methodOfEducation = new List<string>();
            foreach (Word.Cell cell in tableInDoc.Range.Cells){
                string tempGoal = cell.Range.Text.Replace("\v", " ");
                tempGoal = tempGoal.Replace("\r", " ");
                tempGoal = tempGoal.Replace("\a", " ");
                tempGoal = Char.ToUpper(tempGoal[0]) + tempGoal.Substring(1);
                tempGoal = tempGoal.Trim();
                additionalPhraze.Add(tempGoal);
            }
            tableInDoc = docWithGoals.Tables[3];
            foreach (Word.Cell cell in tableInDoc.Range.Cells)
            {
                string tempGoal = cell.Range.Text.Replace("\v", " ");
                tempGoal = tempGoal.Replace("\r", " ");
                tempGoal = tempGoal.Replace("\a", " ");
                tempGoal = Char.ToUpper(tempGoal[0]) + tempGoal.Substring(1);
                tempGoal = tempGoal.Trim();
                placeOfEducations.Add(tempGoal);
            }
            tableInDoc = docWithGoals.Tables[3];
            foreach (Word.Cell cell in tableInDoc.Range.Cells)
            {
                string tempGoal = cell.Range.Text.Replace("\v", " ");
                tempGoal = tempGoal.Replace("\r", " ");
                tempGoal = tempGoal.Replace("\a", " ");
                tempGoal = Char.ToUpper(tempGoal[0]) + tempGoal.Substring(1);
                tempGoal = tempGoal.Trim();
                methodOfEducation.Add(tempGoal);
            }
            FilesAPI.WordAPI.Close(docWithGoals);
            Dictionary<string, List<string>> requirementsForStudent = discipline.RequirementsForStudent;
            this.documentPath = documentPath;
            this.documentPathPlan = documentPathPlan;
            nameDiscipline.Text = discipline.Name;
            numberTopic.Text = topic.Name.Substring(0, topic.Name.IndexOf("«"));
            string temp = topic.Name.Substring(topic.Name.IndexOf("«"));
            topicName.Text = temp.Substring(0,temp.Length);
            numberLesson.Text = lesson.LessonInMaterialSupp;
            lessonName.Text = lesson.ThemeOfLesson;
            foreach(string additionalPh in additionalPhraze)
            {
                additionalPhraze_1.Items.Add(additionalPh);
                additionalPhraze_2.Items.Add(additionalPh);
                additionalPhraze_3.Items.Add(additionalPh);
                additionalPhraze_4.Items.Add(additionalPh);
                additionalPhraze_5.Items.Add(additionalPh);
            }
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
            foreach (String goal in requirementsForStudent["Воспитательные:"])
            {
                string tempGoal = goal.Replace("\v", " ");
                tempGoal = tempGoal.Replace("\r", " ");
                tempGoal = tempGoal.Replace("\a", " ");
                tempGoal = Char.ToUpper(tempGoal[0]) + tempGoal.Substring(1);
                tempGoal = tempGoal.Trim();
                selectGoal_4.Items.Add(tempGoal);
            }
            kind.Text = lesson.Type.Substring(0,lesson.Type.LastIndexOf(' '));
            
            foreach(String pl in placeOfEducations)
            {
                place.Items.Add(pl);
            }
            foreach (String pl in methodOfEducation)
            {
                method.Items.Add(pl);
            }

            intro_combo.Items.Add("Принять рапорт дежурного по взводу");
            intro_combo.Items.Add("Проверить наличие и внешний вид обучаемых, сделать необходимые отметки в журнале учета занятий и воспитательной работы");
            intro_combo.Items.Add("Опросить студентов по заданному для повторения к данному занятию материалу");
            intro_combo.Items.Add("Выслушать дополнения, произвести разбор ответов, отметить студентов хорошо подготовившихся к занятию и объявить оценки");
            intro_combo.Items.Add("Сделать вывод о подготовке взвода к занятию");
            intro_combo.Items.Add("Объявить и дать под запись номер и название темы");
            intro_combo.Items.Add("Объявить об отведенном времени на данную тему");
            intro_combo.Items.Add("Объявить и дать под запись номер и название занятия");
            intro_combo.Items.Add("Объявить об отведенном времени на данное занятие");
            intro_combo.Items.Add("Объявить цель занятия");
            intro_combo.Items.Add("Объявить и дать под запись учебные вопросы занятия");
            intro_combo.Items.Add("Ознакомить с рекомендованной литературой по данному занятию");
            intro_combo.Items.Add("Проинструктировать студентов по мерам безопасности при отработке учебного материала занятия");

            conclusion_combo.Items.Add("Уточнить у обучающихся, есть ли у них вопросы по материалу занятия и ответить на неясные вопросы");
            conclusion_combo.Items.Add("Подвести итоги данного занятия. Сделать выводы – достигло ли занятие своей цели");
            conclusion_combo.Items.Add("Объявить и выставить в журнал учета оценки за ответы, отметить наиболее активных обучающихся и слабо успевающих, указать недостатки, сроки их устранения");
            conclusion_combo.Items.Add("Отметить дисциплину, организованность учебной группы");
            conclusion_combo.Items.Add("Нацелить обучающихся на следующее занятие, дать задание на самостоятельную подготовку и указать литературу");
            hours.Text = lesson.Minutes+" минут";
            
            if (lesson.Type.StartsWith("Лекция")) {
                introLabel.Content = "Вступительная часть:";
            }

            materialSupport.Text = lesson.MaterialSupport;
            literature.Text = lesson.Literature.Trim().Replace("\r", "; ").Replace("  ","").Replace("   ", "");
            for(int i=0;i< lesson.Questions.Count; i++)
            {
                if (i == 0)
                {
                    questionName1.Text = lesson.Questions[i];
                    question1_text.IsEnabled = true;
                    question1_time.IsEnabled = true;

                }
                else if(i ==1)
                {
                    questionName2.Text = lesson.Questions[i];
                    question2_text.IsEnabled = true;
                    question2_time.IsEnabled = true;
                }
                else if (i == 2)
                {
                    questionName3.Text = lesson.Questions[i];
                    question3_text.IsEnabled = true;
                    question3_time.IsEnabled = true;
                }
                else if (i == 3)
                {
                    questionName4.Text = lesson.Questions[i];
                    question4_text.IsEnabled = true;
                    question4_time.IsEnabled = true;
                }
                else if (i == 4)
                {
                    questionName5.Text = lesson.Questions[i];
                    question5_text.IsEnabled = true;
                    question5_time.IsEnabled = true;
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

        private void SelectGoal_4_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            goals_4.Text += selectGoal_4.SelectedItem + "; ";
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (literature.Text.Length==0||place.Text.Length==0||materialSupport.Text.Length==0||intro_time.Text.Length==0||question1_time.Text.Length==0)
            {
                ErrorWindow error = new ErrorWindow();
                error.Show();
            }
            else {
                Dictionary<string, Object> keyValuePairs = new Dictionary<string, object>();
                List<string> goals = goals_1.Text.Replace("; ", ";").Split(';').ToList<string>();
                goals.AddRange(goals_2.Text.Replace("; ", ";").Split(';').ToList<string>());
                goals.AddRange(goals_3.Text.Replace("; ", ";").Split(';').ToList<string>());
                goals.AddRange(goals_4.Text.Replace("; ", ";").Split(';').ToList<string>());
                goals.RemoveAll(x => x.Length.Equals(0));
                //List<string> goals = new List<string>();
                //goals.Add("Требоване 1");
                //goals.Add("Требоване 2");
                //goals.Add("Требоване 3");
                //goals.Add("Требоване 4");
                Dictionary<string, string> questions = new Dictionary<string, string>();
                char separator = '$';
                int sumOfMinInQuestionsOfLesson = 0;
                questions.Add(questionName1.Text.Substring(0), question1_time.Text + " мин"+separator+ question1_text.Text);
                sumOfMinInQuestionsOfLesson += Int32.Parse(question1_time.Text);
                if (question2_text.IsEnabled)
                {
                    questions.Add(questionName2.Text.Substring(0), question2_time.Text + " мин" + separator + question2_text.Text);
                    sumOfMinInQuestionsOfLesson += Int32.Parse(question2_time.Text);
                }
                if (question3_text.IsEnabled)
                {
                    questions.Add(questionName3.Text.Substring(0), question3_time.Text + " мин" + separator + question3_text.Text);
                    sumOfMinInQuestionsOfLesson += Int32.Parse(question3_time.Text);
                }
                if (question4_text.IsEnabled)
                {
                    questions.Add(questionName4.Text.Substring(0), question4_time.Text + " мин" + separator + question4_text.Text);
                    sumOfMinInQuestionsOfLesson += Int32.Parse(question4_time.Text);
                }
                if (question5_text.IsEnabled)
                {
                    questions.Add(questionName5.Text.Substring(0), question5_time.Text + " мин" + separator + question5_text.Text);
                    sumOfMinInQuestionsOfLesson += Int32.Parse(question5_time.Text);
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
                keyValuePairs["{id:intro}"] = intro_time.Text+separator+intro_text.Text;
                keyValuePairs["{id:educationalQuestions}"] = sumOfMinInQuestionsOfLesson;
                keyValuePairs["{id:questions}"] = questions;
                keyValuePairs["{id:conclution}"] = conclusion_time.Text+separator+conclusion_text.Text;
                keyValuePairs["{id:material}"] = materialSupport.Text;
                keyValuePairs["{id:methodical}"] = methodical.Text;
                keyValuePairs["{id:technicalMeans}"] = materialSupport.Text;
                documentPath = documentPath.Replace("//", "");
                documentPathPlan = documentPath.Replace("//", "");
                UpdateDoc update = new UpdateDoc(documentPath,documentPathPlan);
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

        private void AdditionalPhraze_1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            question1_text.Text +="\n"+ additionalPhraze_1.SelectedItem+"\n";
        }

        private void AdditionalPhraze_2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            question2_text.Text += "\n" + additionalPhraze_2.SelectedItem + "\n";
        }

        private void AdditionalPhraze_3_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            question3_text.Text += "\n" + additionalPhraze_3.SelectedItem + "\n";
        }

        private void AdditionalPhraze_4_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            question4_text.Text += "\n" + additionalPhraze_4.SelectedItem + "\n";
        }

        private void AdditionalPhraze_5_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            question5_text.Text += "\n" + additionalPhraze_5.SelectedItem + "\n";
        }

        private void Question1_text_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                question1_text.Text += "\n";
            }
        }

        private void Question2_text_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                question2_text.Text += "\n";
            }
        }

        private void Question3_text_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                question3_text.Text += "\n";
            }
        }

        private void Question4_text_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                question4_text.Text += "\n";
            }
        }

        private void Question5_text_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                question5_text.Text += "\n";
            }
        }

        private void goals_1_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
    }
}
