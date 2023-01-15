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
using Tekla.Structures.Model;
using TeklaTemplate.PackageBuilder;

namespace TeklaTemplate.UI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            if( new Model().GetConnectionStatus())
            {
              Close();
            }
            
            if (System.Diagnostics.Debugger.IsAttached)
            {
                Console.WriteLine("I have the debugger attached!");
                PackageCreator.BuildPackage();
            }
        }
    }
}
