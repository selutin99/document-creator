using System.Windows;
using System.Windows.Input;
using System.Windows.Forms;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Diagnostics;
using System;
using System.ComponentModel;
using System.Threading;
using System;

namespace DocumentCreator
{
    public partial class MainWindow : Window
    {
        private string fileName { get; set; }
        ParseThematicPlan parser;
        public string FolderName { get => folderName; set => folderName = value; }
        internal List<Discipline> Disciplines { get => disciplines; set => disciplines = value; }
        private Dictionary<string, List<string>> requirementsForStudent;
        private List<Discipline> disciplines;

        private string folderName = @"C:\out\";
        private FolderBrowserDialog folderBrowserDialog1;

        public MainWindow()
        {
            InitializeComponent();
            CreateFolder(folderName);
        }
        public static void CreateFolder(string folderPath)
        {
            try
            {
                if (Directory.Exists(folderPath))
                {
                    return;
                }

                DirectoryInfo di = Directory.CreateDirectory(folderPath);
            }
            catch (Exception e)
            {
                Console.WriteLine("Не могу создать папку!");
            }
        }

        private void DownloadButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".doc";
            dlg.Filter = "Word documents (.doc)|*.doc|(.docx)|*.docx";

            dynamic result = dlg.ShowDialog();

            if (result == true)
            {
                DialogWindow dialogWindow = new DialogWindow();
                fileName = dlg.FileName;
                PathToFile.Content = fileName; //вывод в окно имени файла
                //dialogWindow.unswerLabel.Content = dlg.FileName + "\nуспешно загружен!";
                dialogWindow.unswerLabel.Content = "Темплан успешно загружен!";
                dialogWindow.Show();
            }
            CheckEnabledForGenerate();
        }

        //Сгенерировать УМР
        private void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            string sourceDir = System.IO.Path.GetDirectoryName(fileName);
            string backupDir = FolderName;
            string fName = System.IO.Path.GetFileName(fileName);

            //Логика
            ParseWorkPrograming parseWorkPrograming = new ParseWorkPrograming("C:\\programma.docx");
            requirementsForStudent = parseWorkPrograming.ParsePlan();
            parser = new ParseThematicPlan(fileName, FolderName+"//");
            Disciplines = parser.ParseThematicPlanAndCreateDirectories();
            //List<Discipline> disciplines = parser.ParseThematicPlanAndCreateDirectories();
            foreach (Discipline discipline in Disciplines)
                {
                    ComboDisciplines.Items.Add(discipline.Name);
                    Directory.CreateDirectory(parser.GetOutputPath() + discipline.Name);
                    foreach (Topic topic in discipline.Topics)
                    {
                        Directory.CreateDirectory(parser.GetOutputPath() + discipline.Name + "\\" + topic.Name);
                        for (int i = 0; i < topic.Lessons.Count;i++)
                        {
                            Lesson lesson = topic.Lessons[i];
                            //Disciplene disciplineWindow = new Disciplene();
                            //disciplineWindow.NameOfDiscipline.Content = discipline.Name;
                            //disciplineWindow.Theme.Content = topic.Name;
                            //disciplineWindow.LessonType.Content = lesson.Type;
                            //disciplineWindow.ShowDialog();
                            string path = Path.GetFullPath(Path.Combine(System.Reflection.Assembly.GetExecutingAssembly().Location, @"../../../../../Resources/"));
                            string fileName = path + "theme.doc";
                            string outputFileName = parser.GetOutputPath() + discipline.Name + "\\" + topic.Name + "\\" + lesson.Type + ".doc";
                            outputFileName = outputFileName.Replace("//", "\\");
                            File.Copy(@fileName, @outputFileName);
                        }
                    }
                }
            /*ParseThematicPlan parser = new ParseThematicPlan(fileName, folderName);
            parser.ParseThematicPlanAndCreateDirectories();*/
            File.Copy(System.IO.Path.Combine(sourceDir, fName), System.IO.Path.Combine(backupDir, fName), true);
            DialogWindow dialogWindow = new DialogWindow();
            dialogWindow.makeOpenButtonEnabled();
            string folderName = @"C:\out\";
            Process.Start(folderName);
            dialogWindow.unswerLabel.Content = "УМР успешно созданы";
            dialogWindow.Show();   
        }

        private void CheckEnabledForGenerate()
        {
            if (string.IsNullOrEmpty(PathToFile.Content.ToString()))
            {
                GenerateButton.IsEnabled = false;
            }
            else
            {
                GenerateButton.IsEnabled = true;
            }
        }

        private void ComboDisciplines_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            ComboLesson.Items.Clear();
            foreach (Discipline discipline in Disciplines)
            {
                if (discipline.Name.Equals(ComboDisciplines.SelectedItem.ToString()))
                {
                    ComboTheme.Items.Insert(0, "Выберите тему");
                    ComboTheme.SelectedIndex = 0;
                    int count = ComboTheme.Items.Count;
                        for (int i = 1; i < count; i++)
                        {
                            ComboTheme.Items.RemoveAt(1);
                        }
                        //ComboTheme.SelectedIndex = 0;
                        foreach (Topic topic in discipline.Topics)
                        {
                            ComboTheme.Items.Add(topic.Name);
                        }
                }
            }
        }

        private void ComboTheme_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
                ComboLesson.Items.Clear();
                foreach (Discipline discipline in Disciplines)
                {
                    if (discipline.Name.Equals(ComboDisciplines.SelectedItem.ToString()))
                    {
                        foreach (Topic topic in discipline.Topics)
                        {
                            if (topic.Name.Equals(ComboTheme.SelectedItem.ToString()))//Ошибка при повторном выборе дисциплины
                            {
                                for (int i = 0; i < topic.Lessons.Count; i++)
                                {
                                    Lesson lesson = topic.Lessons[i];
                                    ComboLesson.Items.Add(lesson.Type);
                                }
                            }
                        }
                    }
                }
        }
        //ComboBox.Remove();
        private void ComboLesson_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            ChangeButton.IsEnabled = true;
        }

        private void ChangeButton_Click(object sender, RoutedEventArgs e)
        {
            Console.WriteLine(ComboLesson.SelectedIndex);
            string documentPath= parser.GetOutputPath() + ComboDisciplines.SelectedItem + "\\" + ComboTheme.SelectedItem + "\\" + ComboLesson.SelectedItem + ".doc";
            String firstSymb = ComboLesson.SelectedItem.ToString();
            String discipline = ComboDisciplines.SelectedItem.ToString();
            String theme = ComboTheme.SelectedItem.ToString();
            String lesson = ComboLesson.SelectedItem.ToString();
            Discipline selectedDiscipline = Disciplines.Find(x => x.Name.Equals(discipline));
            Topic selectedTopic = selectedDiscipline.Topics.Find(x => x.Name.Equals(theme));
            Lesson selectedLesson = selectedTopic.Lessons.Find(x => x.Type.Equals(lesson));
            firstSymb = firstSymb[0].ToString();
            if (String.Compare(firstSymb, "Л") == 0)
            {
                ChangeWindow change = new ChangeWindow();
                Dictionary<string, Object> keyValuePairs = new Dictionary<string, object>();
                List<string> goals = new List<string>();
                goals.Add("Требоване 1");
                goals.Add("Требоване 2");
                goals.Add("Требоване 3");
                goals.Add("Требоване 4");
                Dictionary<string,string> questions = new Dictionary<string, string>();
                questions.Add("Вопрос фцворфшцапш фнцпшнфцашцф ивфцивфшцпнц шфгвршг црвфш гврфц 1","20 мин");
                questions.Add("Вопрос фцвшгрфшдцгап шФГЦПшгпш гфца 2", "30 мин");
                questions.Add("Вопрос фгцщрашфг апгнпцшфгарфцшгнапг нлфпагцфнапгноцфп ифцгнавпцгфп аифцоври 3", "40 мин");
                keyValuePairs["{id:name}"]="Название дисциплины";
                keyValuePairs["{id:theme}"] = "Тема N1";
                keyValuePairs["{id:themeName}"] = "Название темы";
                keyValuePairs["{id:lesson}"] = "Занятие 1";
                keyValuePairs["{id:lessonName}"] = "Назване занятия";
                keyValuePairs["{id:goal}"] = goals;
                keyValuePairs["{id:kind}"] = "Групповое заянтие";
                keyValuePairs["{id:method}"] = "Метод проведения занятия";
                keyValuePairs["{id:duration}"] = "2 часа";
                keyValuePairs["{id:place}"] = "Плац";
                keyValuePairs["{id:literature}"] = "ЛИТЕРАтура занятия";
                keyValuePairs["{id:intro}"] = "10";
                keyValuePairs["{id:educationalQuestions}"] = "50";
                keyValuePairs["{id:questions}"] = questions;
                keyValuePairs["{id:conclution}"] = "10";
                keyValuePairs["{id:material}"] = "Материалы занятия!!";
                keyValuePairs["{id:additionalLiterature}"] = "Дополнительная литература!!";
                keyValuePairs["{id:technicalMeans}"] = "Технические средства!!";
                UpdateDoc update = new UpdateDoc(documentPath);
                update.updateDoc(keyValuePairs);
                change.initValues(selectedDiscipline, selectedTopic, selectedLesson, requirementsForStudent);
                change.Show();
            }
            else if (String.Compare(firstSymb, "С") == 0)
            {
                //Открытие окна самостоятельных работ
            }
            else if (String.Compare(firstSymb, "Г") == 0)
            {
                //Открытие окна групповых занятий
            }
            else if (String.Compare(firstSymb, "Т") == 0)
            {
                //Открытие окна тренировок 
            }
            else if (String.Compare(firstSymb, "П") == 0)
            {
                //Открытие окна практических занятий
            }
        }
    }
}
