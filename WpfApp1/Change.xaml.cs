using System;
using System.Collections.Generic;
using System.Drawing;
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
            int times = 0;
            var names= txt_name.Text.Split('，');
            var changes = txt_change.Text.Split("+");
            var reasons = txt_reason.Text.Split("，");
            AllGroup ag = new AllGroup();
            
                foreach (var group in ag.Groups)
                {if (times == names.Length) break;
                foreach (var student in group.Students)
                {
                    foreach (var name in names)
                    {
                        if (student.name == name)
                        {
                            for(int i = 0; i < changes.Length; i++) {
                                student.change = float.Parse(changes[i]);
                                student.reason = reasons[i];
                                student.ChangeScore();
                            }
                            times++;
                            break;
                        }
                    }
                    
                }
                }
           
            ag.Ranking();
            ag.Save();

            if (times == names.Length) { MessageBox.Show("修改成功！"); }
            else { MessageBox.Show("该学生不存在！"); }
            txt_change.Clear();
            txt_name.Clear();
            txt_reason.Clear();
        }
        
            
    }
}
