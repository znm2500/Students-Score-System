using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
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
        public int[] duqu_lie = new int[6];
        public int _number = 0, _start_row = 0,mz_row=0,zf_row=0;

        public Exam(string filepath, bool duikang, int type = 0, bool excel = true, bool docx=true,Data data=null)
        {
            InitializeComponent();
            try
            {
                foreach (var i in ag.Groups)
                {
                    students.AddRange(i.Students);
                }
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
                        _number = data._number;
                        duqu_lie[0] = data.dl0;
                        duqu_lie[1] = data.dl1;
                        duqu_lie[2]= data.dl2;
                        duqu_lie[3]= data.dl3;
                        duqu_lie[4] = data.dl4;
                        duqu_lie[5] = data.dl5;
                        mz_row= data.mz_row;
                        zf_row= data.zf_row;
                        _start_row = data._start_row;
                      
                        for (int i = _start_row; i < _start_row + _number; i++)
                        {
                            name = sheet.GetRow(i).GetCell(mz_row).StringCellValue;
                            
                            foreach (var gp in ag.Groups)
                            {
                                foreach (var st in gp.Students)
                                {
                                    if (st.name == name || (st.name.Contains("周怡") && name.Contains("周怡")) || (st.name.Contains("王婷") && name.Contains("王婷")))
                                    {
                                        var alls = sheet.GetRow(i).GetCell(zf_row);
                                        alls.SetCellType(CellType.String);
                                        if (float.TryParse(alls.StringCellValue,out _))
                                        {
                                            st.all_exam_score = float.Parse(alls.StringCellValue);
                                        }
                                        else { st.all_exam_score = 0;
                                            st.valid = false;
                                        }

                                        for (int j = 0; j < 6; j++)
                                        {

                                            sheet.GetRow(i).GetCell(duqu_lie[j]).SetCellType(CellType.String);
                                            single = sheet.GetRow(i).GetCell(duqu_lie[j]).StringCellValue;

                                            if (!float.TryParse(single, out _))
                                            {
                                                st.examscore[j] = 0f;
                                                st.valid = false;
                                            }
                                            else
                                            {
                                                
                                                st.examscore[j] = float.Parse(single);

                                            }

                                        }
                                        break;
                                    }
                                }
                            }

                        }
                        for (int i = 0; i < 6; i++)
                        {
                            students.Sort((b, a) => a.examscore[i].CompareTo(b.examscore[i]));
                            for (int j = 0; j <= 9; j++)
                            {
                                students[j].examrank[i] = j+1;
                                
                            }
                        }
                        foreach (var st in students) st.AfterExam();
                        ag.rankingscore();
                    }
                }
                else if (docx)
                {
                    var word = new XWPFDocument(File.OpenRead(filepath));

                    var sheet = word.Tables[0];
                    string name = "";

                    if (duikang)
                    {

                        float avg, sum = 0, total = 0;
                        string single, nname;
                        foreach (var row in sheet.Rows)
                        {
                            var cells = row.GetTableCells();
                            if (!float.TryParse(cells[4].GetText(), out _)) break;
                            sum += float.Parse(cells[4].GetText());
                            total++;
                        }//求和
                        avg = sum / total;

                        foreach (var row in sheet.Rows)
                        {
                            var cells = row.GetTableCells();
                            single = cells[4].GetText();
                            nname = cells[2].GetText();
                            if (!float.TryParse(single, out _)) break;

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
                        for (int i = 0; i <= sheet.Rows.Count; i++)
                        {
                            var cells = sheet.Rows[i].GetTableCells();
                            name = cells[2].GetText();
                            foreach (var gp in ag.Groups)
                            {
                                foreach (var st in gp.Students)
                                {
                                    if (st.name == name)
                                    {

                                        st.all_exam_score = float.Parse(cells[4].GetText());

                                        for (int j = 0; j < 6; j++)
                                        {
                                            single = cells[j * 3 + 6].GetText();
                                            if (!float.TryParse(cells[4].GetText(), out _)) { st.examrank[j] = 100; }
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
                else
                {

                    var text = File.ReadAllText(filepath);

                    var informations = text.Split('\n');


                    if (duikang)
                    {

                        float avg, sum = 0, total = 0;
                        string single, nname;
                        foreach (var information in informations)
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
                        {
                            if (information.Length == 0) break;
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


                result.ItemsSource = students;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message);
                Close();
            }

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
