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
            openFileDialog.Title = "选择excel文件";
            openFileDialog.Filter = "excel文件|*.xlsx";

            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = false;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.DefaultExt = "";
            if (openFileDialog.ShowDialog() == true)
            {
                Exam si = new Exam(openFileDialog.FileName,true,1);
                si.Owner = this;
                si.ShowDialog();



            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "选择excel文件";
            openFileDialog.Filter = "excel文件|*.xlsx";

            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = false;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.DefaultExt = "";
            if (openFileDialog.ShowDialog() == true)
            {
                Exam si = new Exam(openFileDialog.FileName, false);
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
            openFileDialog.Title = "选择excel文件";
            openFileDialog.Filter = "excel文件|*.xlsx";

            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = false;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.DefaultExt = "";
            if (openFileDialog.ShowDialog() == true)
            {
                Exam si = new Exam(openFileDialog.FileName, true, 0);
                si.Owner = this;
                si.ShowDialog();



            }
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "选择excel文件";
            openFileDialog.Filter = "excel文件|*.xlsx";

            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = false;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.DefaultExt = "";
            if (openFileDialog.ShowDialog() == true)
            {
                Exam si = new Exam(openFileDialog.FileName, true, 2);
                si.Owner = this;
                si.ShowDialog();



            }
        }

        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "选择excel文件";
            openFileDialog.Filter = "excel文件|*.xlsx";

            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = false;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.DefaultExt = "";
            if (openFileDialog.ShowDialog() == true)
            {
                Exam si = new Exam(openFileDialog.FileName, true, 3);
                si.Owner = this;
                si.ShowDialog();



            }
        }
    }
}
