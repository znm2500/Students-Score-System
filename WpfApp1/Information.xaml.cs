using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace WpfApp1
{
    /// <summary>
    /// Information.xaml 的交互逻辑
    /// </summary>
    public partial class Information : Window
    {
        string fp;
        bool duik, excel, doc;
        int typ;
        
        public Information(string filepath, bool duikang, int type = 0, bool excel = true, bool docx = true)
        {
            fp=filepath;
            duik = duikang;
            typ = type;
            this.excel = excel;
            doc = docx;
            InitializeComponent();
            number.Text = "56";
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var ps = new Data();
                     
            ps._start_row = int.Parse(start_line.Text) - 1;
            var s = Regex.Split(rows.Text, @"\s+");
            var ls=new int[6];
            for(int i = 0; i < 6; i++)
            {
                ls[i]= int.Parse(s[i]);
            }
            ps.dl0= ls[0]-1;
            ps.dl1= ls[1]-1;
            ps.dl2= ls[2] - 1;
            ps.dl3= ls[3] - 1;
            ps.dl4= ls[4] - 1;
            ps.dl5= ls[5] - 1;

            ps._number = int.Parse(number.Text);
 
            ps.mz_row = int.Parse(name.Text)-1;
            ps.zf_row=int.Parse(all.Text)-1;
            var ex = new Exam(fp, duik, typ, excel, doc,ps);
            ex.Owner = this;
            ex.ShowDialog();
        }
    }
}
