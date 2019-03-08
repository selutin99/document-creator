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

namespace DocumentCreator
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void DownloadButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".doc";
            dlg.Filter = "Word documents (.doc)|*.doc|(.docx)|*.docx|(.txt)|*.txt";/*.*/

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
                dialogWindow.unswerLabel.Content = dlg.FileName + "\nуспешно загружен!";
                dialogWindow.Show();
            }
        }

            private void ListOfDisciplines_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {
                DownloadButton.IsEnabled = true;
                string documentName = ((Button)ListOfDisciplines.SelectedItem).Content.ToString();
            }
    }
}
