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

namespace WpfApp1
{
    /// <summary>
    /// Change.xaml 的交互逻辑
    /// </summary>
    public partial class Change : Window
    {

        public Change()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            bool ok = false;
            string name = txt_name.Text;
            float change = float.Parse(txt_change.Text);
            string reason = txt_reason.Text;
            AllGroup ag = new AllGroup();
            foreach (var group in ag.Groups)
            {
                foreach (var student in group.Students)
                {
                    if (student.name == name)
                    {
                        student.change = change;
                        student.reason = reason;
                        student.ChangeScore();
                        ok = true;
                        break;
                    }
                }
                if(ok) { break; }
            }
            ag.Ranking();
            ag.Save();

            if (ok) { MessageBox.Show("修改成功！"); }
            else { MessageBox.Show("该学生不存在！"); }
            txt_change.Clear();
            txt_name.Clear();
            txt_reason.Clear();
        }
        
            
    }
}
