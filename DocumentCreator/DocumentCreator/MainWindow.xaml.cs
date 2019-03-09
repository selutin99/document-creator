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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms;
using System.IO;

namespace DocumentCreator
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string documentName;
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
            
            //Nullable<bool> result = dlg.ShowDialog();
            //if (result == true)
            //{
            //    if (dlg.FileName.Length > 0)
            //    {
            //        SelectedFileTextBox.Text = dlg.FileName;
            //        string newXPSDocumentName = String.Concat(System.IO.Path.GetDirectoryName(dlg.FileName), "\\",
            //                       System.IO.Path.GetFileNameWithoutExtension(dlg.FileName), ".xps");

            //        documentViewer1.Document =
            //            ConvertWordDocToXPSDoc(dlg.FileName, newXPSDocumentName).GetFixedDocumentSequence();
            //    }
            //}
            if (result == true)
            {
                DialogWindow dialogWindow = new DialogWindow();
                fileName = dlg.FileName;
                PathToFile.Content = fileName; //вывод в окно имени файла
                dialogWindow.unswerLabel.Content = dlg.FileName + "\nуспешно загружен!";
                dialogWindow.Show();
            }
        }

        private void ListOfDisciplines_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            DownloadButton.IsEnabled = true;
            PathToSaveButton.IsEnabled = true;
            documentName = ((System.Windows.Controls.Button)ListOfDisciplines.SelectedItem).Content.ToString();
        }

        private void PathToSaveButton_Click(object sender, RoutedEventArgs e)
        {
            folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            DialogResult result = folderBrowserDialog1.ShowDialog();           
            folderName = folderBrowserDialog1.SelectedPath;
            SavePath.Content = folderName;
        }

        //Загружает файл в указанную директорию
        private void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            string sourceDir = System.IO.Path.GetDirectoryName(fileName);
            string backupDir = folderName;
            string fName = System.IO.Path.GetFileName(fileName);
            File.Copy(System.IO.Path.Combine(sourceDir, fName), System.IO.Path.Combine(backupDir, fName), true);
            DialogWindow dialogWindow = new DialogWindow();
            dialogWindow.unswerLabel.Content = "Файл успешно сохранен";
            dialogWindow.Show();

        }
    }
}
