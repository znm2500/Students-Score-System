using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
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
    }
}
