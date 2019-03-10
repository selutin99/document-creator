using System.Windows;
using System.Windows.Input;
using System.Windows.Forms;
using System.IO;

namespace DocumentCreator
{
    public partial class MainWindow : Window
    {
        private string fileName { get; set; }
        private string folderName { get; set; }
        private FolderBrowserDialog folderBrowserDialog1;

        public MainWindow()
        {
            InitializeComponent();
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

        private void PathToSaveButton_Click(object sender, RoutedEventArgs e)
        {
            folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            DialogResult result = folderBrowserDialog1.ShowDialog();           
            folderName = folderBrowserDialog1.SelectedPath;
            SavePath.Content = folderName;
            CheckEnabledForGenerate();
        }

        //Сгенерировать УМР
        private void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            string sourceDir = System.IO.Path.GetDirectoryName(fileName);
            string backupDir = folderName;
            string fName = System.IO.Path.GetFileName(fileName);
            //Сюда пихать логику

            File.Copy(System.IO.Path.Combine(sourceDir, fName), System.IO.Path.Combine(backupDir, fName), true);
            DialogWindow dialogWindow = new DialogWindow();
            dialogWindow.unswerLabel.Content = "УМР успешно созданы";
            dialogWindow.Show();   
        }

        private void CheckEnabledForGenerate()
        {
            if (string.IsNullOrEmpty(PathToFile.Content.ToString()) ||
               string.IsNullOrEmpty(SavePath.Content.ToString()))
            {
                GenerateButton.IsEnabled = false;
            }
            else
            {
                GenerateButton.IsEnabled = true;
            }
        }
    }
}
