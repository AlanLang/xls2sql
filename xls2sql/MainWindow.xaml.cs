using System;
using System.Collections.Generic;
using System.Data.OleDb;
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

namespace xls2sql
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            OleDbConnection conn = new OleDbConnection("");
            InitializeComponent();
        }


        private void Msg(string msg)
        {
            messagelog.AppendText(msg + "\n");
            msg += "  [" + DateTime.Now.ToString("yyyyMMddHHmmss") + "]";
            string logPath = AppDomain.CurrentDomain.BaseDirectory + "Log/";
            if (!System.IO.Directory.Exists(logPath))
                System.IO.Directory.CreateDirectory(logPath);
            string logFile = logPath + "Log-" + DateTime.Now.ToString("yyyy-MM-dd") + ".log";
            System.IO.File.AppendAllLines(logFile, new string[] { msg });
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {

        }

        private void Window_Drop(object sender, DragEventArgs e)
        {

        }

        private void ImpFileBtn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void closeBtn_Click(object sender, RoutedEventArgs e)
        {
            Msg("程序退出");
            this.Close();
        }

        private void TextBox_PreviewDragOver(object sender, DragEventArgs e)
        {

        }
    }
}
