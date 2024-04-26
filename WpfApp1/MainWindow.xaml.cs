using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
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
using System.Xml;

namespace WpfApp1
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>

    public class Student
    {
        public float score { get; set; }
        public int group { get; set; }
        public string name { get; set; }
        public float change { get; set; }
        public string reason;
        public float all_exam_score { get; set; }
        public string record { get; set; } = "";
        // public int group;
        public int[] examrank = new int[6];


        public void AfterExam()
        {
            foreach (int i in examrank)
            {
                if (i <= 3) { change += 3; }
                else if (i <= 6) { change += 2; }
                else if (i <= 10) { change += 1; }
            }
            examrank = new int[6];
            reason = "月考";
            return;
        }
        public void ExamAgainst(float avr, float examscore, int type)
        {
            switch (type)
            {
                case 0:
                    if (avr > examscore)
                    {
                        change -= 0.5f;
                    }
                    break;
                case 1:
                    change = 0.5f;
                    break;
                case 2:
                    change = 1f;
                    break;
                case 3:
                    if (avr > examscore)
                    {
                        change -= 1f;
                    }
                    break;
            }
            reason = "对抗赛";

        }
        public void ChangeScore()
        {
            score += change;
            if (change != 0)
            {
                if (reason == "")
            {
                if (change < 0)
                    record = record + change.ToString();
                else
                {
                    record = record + "+" + change.ToString();
                }
            }
            else
            {
                if (change < 0)
                    record = record + change.ToString() + "(" + reason + ")";
                else
                {
                    record = record + "+" + change.ToString() + "(" + reason + ")";
                }
            }
            }

            change = 0f;
            reason = "";
        }
        public Student(float s, string n)
        {
            score = s;
            name = n;
        }
    }
    public class Group
    {
        public int id { get; set; }
        public List<Student> Students = new List<Student>();
        public int rank { get; set; }
        public float totalscore;
        public double total_exam_score;
        public void TotalScore()
        {
            totalscore = 0f;
            foreach (var s in Students)
            {
                totalscore += s.score;
            }
            return;
        }
        public void TotalExamScore()
        {
            total_exam_score = 0;
            foreach (var s in Students)
            {
                total_exam_score += s.all_exam_score;
            }
            return;
        }
        public Group(int i, int r)
        {
            id = i;
            rank = r; 
        }
        
    }
    public class AllGroup
    {
        public List<Group> Groups = new List<Group>();
        public void Ranking()
        {

            foreach (var gp in Groups)
            {
                gp.TotalScore();
            }
            Groups.Sort((b, a) => a.totalscore.CompareTo(b.totalscore));
            for (int i = 0; i < 7; i++)
            {
                Groups[i].rank = i + 1;

            }
        }
        public void rankingscore()
        {
            foreach (var gp in Groups) { gp.TotalExamScore();}
             Groups.Sort((b, a) => a.total_exam_score.CompareTo(b.total_exam_score));
            
                for (int i = 1; i <= 8; i++)
                {
                    if (i == 1 || i == 2)
                    {
                        foreach (var s in Groups[i - 1].Students)
                        {
                            if (s.name == "公共")
                            {
                                s.score = 3;
                                break;
                            }
                        }

                    }
                    else if (i >= 3 && i <= 5)
                    {
                        foreach (var s in Groups[i - 1].Students)
                        {
                            if (s.name == "公共")
                            {
                                s.score = 2;
                                break;
                            }
                        }

                    }
                    else
                    {
                        foreach (var s in Groups[i-1].Students)
                        {
                            if (s.name == "公共")
                            {
                                s.score = 1;
                                break;
                            }
                        }

                    }
                }
            

        
        }
        public AllGroup()
        {
            XmlDocument xml = new XmlDocument();
            xml.Load("Groups.xml");
            XmlNode node = xml.SelectSingleNode("Groups");
            XmlNodeList groups = node.ChildNodes;
            int i = -1;
            foreach (XmlNode group in groups)
            {
                XmlElement xegroup = (XmlElement)group;
                int id = int.Parse(xegroup.GetAttribute("id"));
                int rank = int.Parse(xegroup.GetAttribute("rank"));
                var g = new Group(id, rank);
                Groups.Add(g);
                i++;
                XmlNodeList students = group.ChildNodes;
                foreach (XmlNode student in students)
                {
                    XmlElement xestudent = (XmlElement)student;
                    float score = float.Parse(xestudent.GetAttribute("score"));
                    string name = xestudent.GetAttribute("name");

                    Student st = new Student(score, name);
                    st.group = id;
                    st.record = xestudent.GetAttribute("record");
                    Groups[i].Students.Add(st);
                    Groups[i].TotalScore();
                }
                

            }
            Ranking();

        }
        public void Save()
        {
            
            XmlDocument xmlDoc = new XmlDocument();
            var root = xmlDoc.CreateElement("Groups");
            xmlDoc.AppendChild(root);
            foreach (var group in Groups)
            {
                string id = group.id.ToString();
                string rank = group.rank.ToString();
                var gro = xmlDoc.CreateElement("Group");
                gro.SetAttribute("rank", rank);
                gro.SetAttribute("id", id);
                root.AppendChild(gro);
                foreach (var s in group.Students)
                {
                    string score = s.score.ToString();
                    var st = xmlDoc.CreateElement("Student");
                    st.SetAttribute("name", s.name);
                    st.SetAttribute("score", score);
                    st.SetAttribute("record", s.record);
                    gro.AppendChild(st);


                }

            }
            xmlDoc.Save("Groups.xml");
        }


    }

    public partial class MainWindow : Window
    {
        AllGroup ag;
        public AllGroup Load()
        {
            AllGroup al = new AllGroup();
            List<Student> students = new List<Student>();
            
            foreach (var i in al.Groups)
            {
                students.AddRange(i.Students);
            }
            ranking.ItemsSource = students;
            return al;
        }
        public MainWindow()
        {

            InitializeComponent();

            try
            {
                ag = Load();
            }
            catch (Exception ex) { }
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
                StudentImport si = new StudentImport(openFileDialog.FileName);
                si.Owner = this;
                si.ShowDialog();

                Load();


            }

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Import im = new Import();
            im.Owner = this;
            im.ShowDialog();
            Load();

        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            var dialogResult = MessageBox.Show("确认刷新吗？", "提示", MessageBoxButton.OKCancel);
            if (dialogResult == MessageBoxResult.OK)
            {
                foreach (var g in ag.Groups)
                {
                    g.rank = 1;
                    foreach (var s in g.Students)
                    {
                        s.score = 0;
                        s.record = "";
                    }
                }
                ag.Save();
                MessageBox.Show("刷新成功！");
            }
            Load();
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            ag = Load();
            ag.Ranking();
            string str = String.Format($"第一名：第{ag.Groups[0].id}组 {ag.Groups[0].totalscore}分\n第二名：第{ag.Groups[1].id}组 {ag.Groups[1].totalscore}分\n第三名：第{ag.Groups[2].id}组 {ag.Groups[2].totalscore}分\n第四名：第{ag.Groups[3].id}组 {ag.Groups[3].totalscore}分\n第五名：第{ag.Groups[4].id}组 {ag.Groups[4].totalscore}分\n第六名：第{ag.Groups[5].id}组 {ag.Groups[5].totalscore}分\n第七名：第{ag.Groups[6].id}组 {ag.Groups[6].totalscore}分\n第八名：第{ag.Groups[7].id}组 {ag.Groups[7].totalscore}分");
            MessageBox.Show(str);
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            ag = Load();
            ag.Ranking();
            var workbook = new XSSFWorkbook();
            var sheet=workbook.CreateSheet();
            sheet.SetColumnWidth(0, 20 * 256);
            var head=sheet.CreateRow(0);
            var head_cell=head.CreateCell(0);
            head_cell.SetCellValue("小组");
            var head_cell1 = head.CreateCell(1);
            head_cell1.SetCellValue("名字");
            var head_cell2 = head.CreateCell(2);
            head_cell2.SetCellValue("分数");
            var head_cell3 = head.CreateCell(3);
            head_cell3.SetCellValue("加分记录");
            List<IRow>rows = new List<IRow>();
            int i = 0;
            for(var index= 1; index <= ag.Groups.Count;index++)
            {
                var temp = i;
                for(var m=0;m< ag.Groups[index-1].Students.Count;m++)
                {
                    //MessageBox.Show(index.ToString());
                    rows.Add(sheet.CreateRow(++i));
                    var cells=new List<ICell>();
                    for(var j=0;j<4;j++)
                        cells.Add(rows[i - 1].CreateCell(j));
              
                    if (m==0) {

                        cells[0].SetCellValue(string.Format($"第{index}组（{ag.Groups[index - 1].totalscore}分）"));
                    }
                    cells[1].SetCellValue(ag.Groups[index - 1].Students[m].name.ToString());
                    cells[2].SetCellValue(ag.Groups[index - 1].Students[m].score.ToString());
                    cells[3].SetCellValue(ag.Groups[index - 1].Students[m].record);
                }
                sheet.AddMergedRegion(new CellRangeAddress(temp+1, i, 0, 0));
                temp = i;
            }
            
            OpenFolderDialog dialog = new OpenFolderDialog();
            dialog.Multiselect = true;
            dialog.Title = "请选择文件夹";
            if (dialog.ShowDialog()==true)
            {
                FileStream file = new FileStream(dialog.FolderName+"\\分数.xlsx", FileMode.Create);
                workbook.Write(file);
                MessageBox.Show("文件已导出！");

            }

           
        }
    }
}
