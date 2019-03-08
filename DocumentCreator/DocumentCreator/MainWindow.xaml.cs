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
using System.IO;

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
            test();
        }

        private void label_MouseDown(object sender, MouseButtonEventArgs e)
        {
            string path = System.IO.Path.GetFullPath(System.IO.Path.Combine(System.Reflection.Assembly.GetExecutingAssembly().Location, @"../../../../../Resources/"));
            string fullPath = path + "plane.doc";

            ParseThematicPlan parser = new ParseThematicPlan(fullPath);
        }
        private void test()
        {
            string path = System.IO.Path.GetFullPath(System.IO.Path.Combine(System.Reflection.Assembly.GetExecutingAssembly().Location, @"../../../../../Resources/"));
            string fullPath = path + "plane.doc";
            ParseThematicPlan parser = new ParseThematicPlan(fullPath);
            parser.LogicForParseWordAndSave();
        }
    }
}
