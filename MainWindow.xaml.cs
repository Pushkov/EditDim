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

namespace EditDim3
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }


        public void onShow()
        {
            this.Show();
        }

        public void onHide()
        {
            this.Hide();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void EXIT_Click(object sender, RoutedEventArgs e)
        {
            App.Current.Shutdown();
        }
    }
}
