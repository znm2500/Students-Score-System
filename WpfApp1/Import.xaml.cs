using Microsoft.Win32;
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
using System.Windows.Shapes;
using NPOI.XSSF.UserModel;
using System.Xml;
using System.Diagnostics;
namespace WpfApp1
{
    /// <summary>
    /// Import.xaml 的交互逻辑
    /// </summary>
    public partial class Import : Window
    {
        public Import()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "选择成绩文件";
            openFileDialog.Filter = "成绩|*.xlsx;*.txt;*.docx";

            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = false;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.DefaultExt = "";
            if (openFileDialog.ShowDialog() == true)
            {
                var s = openFileDialog.FileName.Split('.');
               
                Exam si = new Exam(openFileDialog.FileName, true,1, s[1] == "xlsx", s[1] == "docx");
                si.Owner = this;
                si.ShowDialog();



            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "选择成绩文件";
            openFileDialog.Filter = "成绩|*.xlsx;*.txt;*.docx";

            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = false;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.DefaultExt = "";
            if (openFileDialog.ShowDialog() == true)
            {
                var s = openFileDialog.FileName.Split('.');

                Exam si = new Exam(openFileDialog.FileName, false, 0, s[1] == "xlsx", s[1] == "docx");
                si.Owner = this;
                si.ShowDialog();



            }
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            Change cg=new Change();
            cg.Owner= this;
            cg.ShowDialog();
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "选择成绩文件";
            openFileDialog.Filter = "成绩|*.xlsx;*.txt;*.docx";

            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = false;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.DefaultExt = "";
            if (openFileDialog.ShowDialog() == true)
            {
                var s = openFileDialog.FileName.Split('.');

                Exam si = new Exam(openFileDialog.FileName, true, 0, s[1] == "xlsx", s[1] == "docx");
                si.Owner = this;
                si.ShowDialog();



            }
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "选择成绩文件";
            openFileDialog.Filter = "成绩|*.xlsx;*.txt;*.docx";

            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = false;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.DefaultExt = "";
            if (openFileDialog.ShowDialog() == true)
            {
                var s = openFileDialog.FileName.Split('.');

                Exam si = new Exam(openFileDialog.FileName, true, 2, s[1] == "xlsx", s[1] == "docx");
                si.Owner = this;
                si.ShowDialog();



            }
        }

        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "选择成绩文件";
            openFileDialog.Filter = "成绩|*.xlsx;*.txt;*.docx";

            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = false;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.DefaultExt = "";
            if (openFileDialog.ShowDialog() == true)
            {
                var s = openFileDialog.FileName.Split('.');

                Exam si = new Exam(openFileDialog.FileName, true, 3, s[1] == "xlsx", s[1]=="docx");
                si.Owner = this;
                si.ShowDialog();



            }
        }
    }
}
