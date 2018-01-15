using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
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
            Msg("****************");
            Msg("请选择或拖入安装包");
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
            this.DragMove();
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            string filepath = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
            if (System.IO.Path.GetExtension(filepath) != "xls" && System.IO.Path.GetExtension(filepath) != "xlsx")
            {
                Msg("只允许上传Excel文件!");
                return;
            }
            MakeSql(filepath);
        }

        private void ImpFileBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel文件(*.xlsx;*.xls)|*.xlsx;*.xls";
            dialog.Multiselect = false;
            if (dialog.ShowDialog() == true)
            {
                MakeSql(dialog.FileName);
            }
        }

        private void closeBtn_Click(object sender, RoutedEventArgs e)
        {
            Msg("程序退出");
            Application.Current.Shutdown();
        }

        private void TextBox_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.Copy;
            e.Handled = true;
        }

        protected void MakeSql(string filepath)
        {
            DataTable dt = new DataTable();
            try
            {
                ExcelHelper exl = new xls2sql.ExcelHelper(filepath);
                dt = exl.GetSheetTable(0);
                string TableName = System.IO.Path.GetFileNameWithoutExtension(filepath);
                List<string> Names = new List<string>();
                List<string> Sqls = new List<string>();
                foreach (DataColumn dc in dt.Columns)
                {
                    Names.Add(dc.ColumnName);
                }
                int index = 1;
                int errcount = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    List<string> Values = new List<string>();
                    foreach (var item in Names)
                    {
                        Values.Add($"'{dr[item].ToString().Trim()}'");
                    }
                    if (Names.Count != Values.Count)
                    {
                        Msg($"err:第{index}行，字段名和值的数量不匹配，此行跳过");
                        errcount++;
                    }
                    string sql = $"INSERT INTO {TableName} ({string.Join<string>(",", Names)}) VALUES ({string.Join<string>(",", Values)});\n";
                    Sqls.Add(sql);
                    Msg($"成功解析第{index}行");
                    index++;
                }
                string msg = errcount == 0?"全部成功":$"中失败了{errcount}行";
                Msg($"全部解析完毕，共计{index}行,{msg}.");
                SqlShow sqlshow = new SqlShow();
                foreach (var item in Sqls)
                {
                    sqlshow.messagelog.AppendText(item);
                }
                sqlshow.Show();
            }
            catch (Exception ex)
            {
                Msg("打开excel异常："+ex.Message);
            }
        }
    }
}
