using System.Windows;
using System.Windows.Input;

namespace DocumentCreator
{
    public partial class MainWindow : Window
    {

        public MainWindow()
        {
            InitializeComponent();
        }

        private void label_MouseDown(object sender, MouseButtonEventArgs e)
        {
            string path = System.IO.Path.GetFullPath(System.IO.Path.Combine(System.Reflection.Assembly.GetExecutingAssembly().Location, @"../../../../../Resources/"));
            string fullPath = path + "plane.doc";
            string output = path + "output//";

            ParseThematicPlan parser = new ParseThematicPlan(fullPath, output);
            parser.LogicForParseWordAndSave();
        }
    }
}
