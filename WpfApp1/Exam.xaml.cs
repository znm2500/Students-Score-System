using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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
using System.Xml.Linq;

namespace WpfApp1
{
    /// <summary>
    /// Exam.xaml 的交互逻辑
    /// </summary>
    public partial class Exam : Window
    {
        AllGroup ag = new AllGroup();
        List<Student> students = new List<Student>();
        public Exam(string filepath, bool duikang, int type = 0, bool excel = true)
        {
            InitializeComponent();
            if (excel)
            {
                var book = new XSSFWorkbook(filepath);
                var sheet = book.GetSheetAt(0);
                string name = "";


                if (duikang)
                {

                    float avg, sum = 0, total = 0;
                    string single, nname;
                    for (int i = 2; i <= sheet.LastRowNum; i++)
                    {
                        if (sheet.GetRow(i).GetCell(5).StringCellValue == "缺考") break;
                        sum += float.Parse(sheet.GetRow(i).GetCell(5).StringCellValue);
                        total++;
                    }//求和
                    avg = sum / total;

                    for (int i = 2; i <= sheet.LastRowNum; i++)
                    {

                        single = sheet.GetRow(i).GetCell(5).StringCellValue;
                        nname = sheet.GetRow(i).GetCell(1).StringCellValue;
                        if (single == "缺考") break;

                        {
                            foreach (var gp in ag.Groups)
                            {
                                foreach (var st in gp.Students)
                                {
                                    if (st.name == nname)
                                    {

                                        st.ExamAgainst(avg, float.Parse(single), type);

                                    }
                                }
                            }
                        }
                    }
                }


                else
                {
                    string single;
                    for (int i = 2; i <= sheet.LastRowNum; i++)
                    {
                        name = sheet.GetRow(i).GetCell(1).StringCellValue;
                        foreach (var gp in ag.Groups)
                        {
                            foreach (var st in gp.Students)
                            {
                                if (st.name == name)
                                {

                                    st.all_exam_score = float.Parse(sheet.GetRow(i).GetCell(5).StringCellValue);

                                    for (int j = 0; j < 6; j++)
                                    {
                                        single = sheet.GetRow(i).GetCell(10 + j * 3).StringCellValue;
                                        if (single == "-") { st.examrank[j] = 100; }
                                        else
                                        {
                                            st.examrank[j] = int.Parse(sheet.GetRow(i).GetCell(10 + j * 3).StringCellValue);
                                        }
                                    }
                                    st.AfterExam();
                                    break;
                                }
                            }
                        }
                    }

                    ag.rankingscore();
                }
            }
            else {
                var text = File.ReadAllText(filepath);
                var informations = text.Split('\n');


                if (duikang)
                {

                    float avg, sum = 0, total = 0;
                    string single, nname;
                    foreach(var information in informations)
                    {
                        if (information.Length == 0) break;
                        var s = Regex.Split(information, "\\s+", RegexOptions.IgnoreCase);
                        sum += float.Parse(s[4]);
                    }
                    total = informations.Length;
                    avg = sum / total;
                    foreach (var information in informations)
                    {
                        if (information.Length == 0) break;
                        var s = Regex.Split(information, "\\s+", RegexOptions.IgnoreCase);
                        nname = s[2];
                        single = s[4];
                        foreach (var gp in ag.Groups)
                            {
                                foreach (var st in gp.Students)
                                {
                                    if (st.name == nname)
                                    {

                                        st.ExamAgainst(avg, float.Parse(single), type);

                                    }
                                }
                            }
                    }
                   
                }

            
                else
                {
                    
                        string single, nname;
                        foreach (var information in informations)
                        {if (information.Length == 0) break;
                            var s = Regex.Split(information, "\\s+", RegexOptions.IgnoreCase);
                            nname = s[2];
                            foreach (var gp in ag.Groups)
                            {
                                foreach (var st in gp.Students)
                                {
                                    if (st.name == nname)
                                    {

                                        st.all_exam_score = float.Parse(s[4]);

                                        for (int j = 0; j < 6; j++)
                                        {
                                            single = s[6 + j * 3];
                                            if (single == "-") { st.examrank[j] = 100; }
                                            else
                                            {
                                                st.examrank[j] = int.Parse(single);
                                            }
                                        }
                                        st.AfterExam();
                                        break;
                                    }
                                }
                            }
                        }


                        ag.rankingscore();
                    
                    
                }
            }
            foreach (var i in ag.Groups)
            {
                students.AddRange(i.Students);
            }
            result.ItemsSource = students;


        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            foreach (var s in students)
            {
                s.ChangeScore();
            }
            ag.Save();
            MessageBox.Show("分数修改完毕！");
            Close();
        }
    }
}
