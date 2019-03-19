using System;
using System.Collections.Generic;
using System.Diagnostics;
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
    /// Логика взаимодействия для DialogWindow.xaml
    /// </summary>
    public partial class DialogWindow : Window
    {
        public DialogWindow()
        {
            InitializeComponent();

            openFolderButton.IsEnabled = false;
            openFolderButton.Opacity = 0;
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void OpenFolderButton_Click(object sender, RoutedEventArgs e)
        {
            string folderName = @"C:\out\";
            Process.Start(folderName);
        }

        public void makeOpenButtonEnabled()
        {
            openFolderButton.IsEnabled = true;
            openFolderButton.Opacity = 100;
        }
    }
}
