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
using System.IO;
using System.Xml.Linq;
using System.Xml;
namespace WpfApp1
{
    /// <summary>
    /// StudentImport.xaml 的交互逻辑
    /// </summary>
    public partial class StudentImport : Window
    {
        List<Student> sts;
        int groupnum=0;
        
        public List<Student> ExcelRead(string filepath)
        {

            var book = new XSSFWorkbook(filepath);
            var sheet = book.GetSheetAt(0);
            
            List<Student> students = new List<Student>();
            string s = "";
            for (int i = 2; i <= sheet.LastRowNum; i++)
            {
                var cell = sheet.GetRow(i).GetCell(1);
                s = cell.StringCellValue;
                var st = new Student(0f, s);
                st.group =int.Parse( sheet.GetRow(i).GetCell(2).StringCellValue);
                groupnum=(st.group>groupnum)?st.group:groupnum;
                students.Add(st);

            }

            return students;

        }
        public StudentImport(string filepath)
        {

            InitializeComponent();
            
           
                sts = ExcelRead(filepath);
                
                List<Group> gps = new List<Group>();
            
                for (int i = 0; i < groupnum; i++)
                {
                    gps.Add(new Group(i + 1, 0));
                }
                students.ItemsSource = sts;
                
           
            
        }
        

        private void Button_Click(object sender, RoutedEventArgs e)
        {
           XmlDocument xmlDoc = new XmlDocument();
            var root = xmlDoc.CreateElement("Groups");
            xmlDoc.AppendChild(root);
            for (int i = 1; i <= groupnum; i++)
            {string n=i.ToString();
                var gro = xmlDoc.CreateElement("Group");
                gro.SetAttribute("rank", "1");
                gro.SetAttribute("id", n);
                root.AppendChild(gro);
                var pu = xmlDoc.CreateElement("Student");
                pu.SetAttribute("name", "公共");
                pu.SetAttribute("score", "0");
                pu.SetAttribute("record", "");
                gro.AppendChild(pu);
                foreach (var s in sts)
                {
                    if (i  == s.group)
                    {
                        var st = xmlDoc.CreateElement("Student");
                        st.SetAttribute("name",s.name);
                        st.SetAttribute("score", "0");
                        st.SetAttribute("record", "");
                        gro.AppendChild(st);
                    }

                }

            }
            xmlDoc.Save("Groups.xml");
            MessageBox.Show("数据读取完毕");
            Close();

           
            
            
            
        }
    }
}
