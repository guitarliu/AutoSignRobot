using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;


namespace AutoSignRobot
{
    /// <summary>
    /// TrainTimeWork.xaml 的交互逻辑
    /// </summary>
    public partial class TrainTimeWork : UserControl
    {
        public TrainTimeWork()
        {
            InitializeComponent();

            // Add Details to ListView
            StreamReader sr = new StreamReader(AppDomain.CurrentDomain.BaseDirectory + @"..\\..\\DataResources\\PersonnelList.db");
            string[] lines = sr.ReadToEnd().Split('\n');
            sr.Close();

            List<string> cmbxProfession = new List<string> { };
            List<string> cmbxDepartment = new List<string> { };

            foreach (string line in lines)
            {
                if (line.Trim() == string.Empty)
                {
                    continue;
                }

                cmbxDepartment.Add(line.Split('\t')[1]);
                cmbxProfession.Add(line.Split('\t')[2]);

                AttendeesInfo attdInfo = new AttendeesInfo
                {
                    AttendeesName = line.Split('\t')[0],
                    AttendeesDepartment = line.Split('\t')[1],
                    AttendeesProfession = line.Split('\t')[2],
                    AttendeesTelePhone = line.Split('\t')[3].Trim('\n', '\r')
                };

                LvePeopleInfo.Items.Add(attdInfo);
            }

            string[] newCmbxDepartment = cmbxDepartment.Distinct().ToList().ToArray();
            string[] newCbxProfession = cmbxProfession.Distinct().ToList().ToArray();
            CbxProfession.ItemsSource = newCbxProfession;
            CbxDepartment.ItemsSource = newCmbxDepartment;
        }

        public class AttendeesInfo
        {
            public string AttendeesName { get; set; }
            public string AttendeesDepartment { get; set; }
            public string AttendeesProfession { get; set; }
            public string AttendeesTelePhone { get; set; }
        }

        public void ListView_Update()
        {
            // Clean ListView to Zero
            LvePeopleInfo.Items.Clear();

            // Refresh ListView 
            LvePeopleInfo.Items.Refresh();

            // Add Details to ListView
            StreamReader sr = new StreamReader(AppDomain.CurrentDomain.BaseDirectory + @"..\\..\\DataResources\\PersonnelList.db");
            string[] lines = sr.ReadToEnd().Split('\n');
            sr.Close();

            foreach (string line in lines)
            {
                if (line.Trim() == string.Empty)
                {
                    continue;
                }
                else if ((CbxProfession.SelectedIndex == -1) & (CbxDepartment.SelectedIndex != -1))
                {
                    if (CbxDepartment.SelectedItem.ToString() == line.Split('\t')[1])
                    {
                        AttendeesInfo attdInfo = new AttendeesInfo
                        {
                            AttendeesName = line.Split('\t')[0],
                            AttendeesDepartment = line.Split('\t')[1],
                            AttendeesProfession = line.Split('\t')[2],
                            AttendeesTelePhone = line.Split('\t')[3].Trim('\n', '\r')
                        };

                        LvePeopleInfo.Items.Add(attdInfo);
                    }
                }
                else if ((CbxProfession.SelectedIndex != -1) & (CbxDepartment.SelectedIndex == -1))
                {
                    if (CbxProfession.SelectedItem.ToString() == line.Split('\t')[2])
                    {
                        AttendeesInfo attdInfo = new AttendeesInfo
                        {
                            AttendeesName = line.Split('\t')[0],
                            AttendeesDepartment = line.Split('\t')[1],
                            AttendeesProfession = line.Split('\t')[2],
                            AttendeesTelePhone = line.Split('\t')[3].Trim('\n', '\r')
                        };

                        LvePeopleInfo.Items.Add(attdInfo);
                    }
                }
                else if ((CbxProfession.SelectedIndex != -1) & (CbxDepartment.SelectedIndex != -1))
                {
                    if ((CbxProfession.SelectedItem.ToString() == line.Split('\t')[2]) & (CbxDepartment.SelectedItem.ToString() == line.Split('\t')[1]))
                    {
                        AttendeesInfo attdInfo = new AttendeesInfo
                        {
                            AttendeesName = line.Split('\t')[0],
                            AttendeesDepartment = line.Split('\t')[1],
                            AttendeesProfession = line.Split('\t')[2],
                            AttendeesTelePhone = line.Split('\t')[3].Trim('\n', '\r')
                        };

                        LvePeopleInfo.Items.Add(attdInfo);
                    }
                }
                else
                {
                    AttendeesInfo attdInfo = new AttendeesInfo
                    {
                        AttendeesName = line.Split('\t')[0],
                        AttendeesDepartment = line.Split('\t')[1],
                        AttendeesProfession = line.Split('\t')[2],
                        AttendeesTelePhone = line.Split('\t')[3].Trim('\n', '\r')
                    };

                    LvePeopleInfo.Items.Add(attdInfo);
                }
            }
        }

        public void WriteToExcel()
        {
            // The Add method has four reference parameters, all of which are
            // optional. Visual C# allows you to omit arguments for them if
            // the default values are what you want.
            object omissing = Missing.Value;
            string DateResult = "";

            if ((!DpStartDate.SelectedDate.HasValue) | (!DpEndDate.SelectedDate.HasValue) | (TbxStartTime.Text == "") | (TbxEndTime.Text == ""))
            {
                MessageBox.Show("请选择日期及时间！", "错误");
            }
            else
            {
                if (DpStartDate.SelectedDate.Value == DpEndDate.SelectedDate.Value)
                {
                    DateResult = $"{DpStartDate.SelectedDate.Value.ToString("D")}" + " " +
                        $"{TbxStartTime.Text}~{TbxEndTime.Text}";
                }
                else
                {
                    DateResult = $"{DpStartDate.SelectedDate.Value.ToString("D")}{TbxStartTime.Text}—" +
                        $"{DpEndDate.SelectedDate.Value.ToString("D")}{TbxEndTime.Text}";
                }

                // Start Excel and get Application object
                var excelApp = new Excel.Application();
                excelApp.Visible = true;

                // Get a new workbook
                Excel._Workbook excelWB = excelApp.Workbooks.Add(omissing);
                Excel._Worksheet excelSheet = excelWB.ActiveSheet;

                // Merge First five rows
                for (int i = 1; i <= 5; i++)
                {
                    excelSheet.Range[$"A{i}:F{i}"].Merge(omissing);
                }
                for (int i = 1; i <= 5; i++)
                {
                    int ftSize = 0;
                    if (i == 1)
                    {
                        ftSize = 20;
                    }
                    else
                    {
                        ftSize = 16;
                    }
                    excelSheet.Range[$"A{i}:F{i}"].Font.Name = "宋体";
                    excelSheet.Range[$"A{i}:F{i}"].Font.Size = ftSize;
                    excelSheet.Range[$"A{i}:F{i}"].Font.Bold = true;
                }
                // Write First Five Rows' Title
                excelSheet.Range[$"A{1}"].Value2 = $"{TbxTitle.Text}";
                excelSheet.Range[$"A{2}"].Value2 = $"培训名称：{TbxPtcpName.Text}";
                excelSheet.Range[$"A{3}"].Value2 = $"培训讲师：{TbxTalker.Text}";
                excelSheet.Range[$"A{4}"].Value2 = $"培训时间：{DateResult}";
                excelSheet.Range[$"A{5}"].Value2 = $"参加培训部门：{TbxPtcpDepartment.Text}";

                // Set Excel cells center
                for (int i = 1; i <= 5; i++)
                {
                    if (i == 1)
                    {
                        excelSheet.Range[$"A{i}"].HorizontalAlignment = Excel.Constants.xlCenter; //Horizontal Center
                        excelSheet.Range[$"A{i}"].VerticalAlignment = Excel.Constants.xlCenter; //Vertical Center
                    }
                    else
                    {
                        excelSheet.Range[$"A{i}"].HorizontalAlignment = Excel.Constants.xlLeft; //Horizontal Left
                        excelSheet.Range[$"A{i}"].VerticalAlignment = Excel.Constants.xlCenter; //Vertical Center
                    }
                }

                excelSheet.Range["A6:F6"].RowHeight = 9;

                for (int i = 7; i <= 27; i++)
                {
                    int ftSize = 0;
                    if (i == 7) { ftSize = 16; }
                    else { ftSize = 12; }
                    // Set Font Size From A7:F7 to A27:F27
                    excelSheet.Range[$"A{i}:F{i}"].Font.Name = "宋体";
                    excelSheet.Range[$"A{i}:F{i}"].Font.Size = ftSize;
                    excelSheet.Range[$"A{i}:F{i}"].Font.Bold = true;

                    excelSheet.Range[$"A{i}:F{i}"].RowHeight = 26.25;
                    excelSheet.Range[$"A{i}:F{i}"].Borders.LineStyle = Excel.XlBordersIndex.xlEdgeLeft;
                    excelSheet.Range[$"A{i}:F{i}"].Borders.LineStyle = Excel.XlBordersIndex.xlEdgeRight;
                    excelSheet.Range[$"A{i}:F{i}"].Borders.LineStyle = Excel.XlBordersIndex.xlInsideVertical;
                    excelSheet.Range[$"A{i}:F{i}"].Borders.LineStyle = Excel.XlBordersIndex.xlEdgeBottom;
                    excelSheet.Range[$"A{i}:F{i}"].Borders.LineStyle = Excel.XlBordersIndex.xlEdgeTop;
                    excelSheet.Range[$"A{i}:F{i}"].Borders.LineStyle = Excel.XlBorderWeight.xlThin;
                    excelSheet.Range[$"A{i}:F{i}"].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    excelSheet.Range[$"A{i}:F{i}"].HorizontalAlignment = Excel.Constants.xlCenter; // Horizontal Center
                    excelSheet.Range[$"A{i}:F{i}"].VerticalAlignment = Excel.Constants.xlCenter; // Vertical Center
                }

                // Set Excel Columns Width for Behind First Six Rows
                Dictionary<string, double> columnWidthDict = new Dictionary<string, double>
                {
                    {"A7", 6.75},
                    {"B7", 17.5},
                    {"C7", 17.5},
                    {"D7", 6.75},
                    {"E7", 17.5},
                    {"F7", 17.5},
                };

                // Set Excel Columns' Width of A7:F7
                foreach (KeyValuePair<string, double> kvp in columnWidthDict)
                {
                    excelSheet.Range[$"{kvp.Key}"].HorizontalAlignment = Excel.Constants.xlCenter; //Horizontal Center
                    excelSheet.Range[$"{kvp.Key}"].VerticalAlignment = Excel.Constants.xlCenter; //Vertical Center
                    excelSheet.Range[$"{kvp.Key}"].ColumnWidth = kvp.Value;
                }

                // Set Excel SignTable Title List
                List<string> tableTitle = new List<string>
                {
                "序号",
                "签名",
                "部门",
                };

                for (int i = 1; i < 4; i++)
                {
                    excelSheet.Cells[7, i] = tableTitle[i - 1];
                }
                for (int i = 4; i < 7; i++)
                {
                    excelSheet.Cells[7, i] = tableTitle[i - 4];
                }

                for (int i = 8; i < 28; i++)
                {
                    excelSheet.Cells[i, 1] = i - 7;
                    excelSheet.Cells[i, 4] = i + 13;
                }

                // Write Pictures to Excel
                // According to ListView Selected Items
                for (int i = 0; i < LvePeopleInfo.SelectedItems.Count; i++)
                {
                    if (i <= 12) // i + 8 <= 20 => i <= 12
                    {
                        try
                        {
                            string signName = ((AttendeesInfo)LvePeopleInfo.SelectedItems[i]).AttendeesName;
                            string signNamePath = AppDomain.CurrentDomain.BaseDirectory + $@"..\\..\\DataResources\\SignImages\\{signName}姓名.svg"; ; // SignName Image Path

                            string signDepartmentPath = AppDomain.CurrentDomain.BaseDirectory + $@"..\\..\\DataResources\\SignImages\\{signName}部门.svg"; // SignDepartment Image Path

                            var LinkToFile = Microsoft.Office.Core.MsoTriState.msoFalse;
                            var SaveWithDocument = Microsoft.Office.Core.MsoTriState.msoTrue;
                            Excel.Range cellNameRange = excelSheet.Cells[i + 8, 2]; //Choose cell to Insert SignName Image
                            Pictures pics = (Pictures)excelSheet.Pictures(omissing);
                            float picsSignNameWidth = 45;
                            float picsSignNameHeight = 25.714f;
                            double picsSignNameLeft = cellNameRange.Left + picsSignNameWidth / 1.5;
                            double picsSignNameTop = cellNameRange.Top + 1;

                            excelSheet.Shapes.AddPicture(signNamePath, LinkToFile, SaveWithDocument, (float)picsSignNameLeft, (float)picsSignNameTop, picsSignNameWidth, picsSignNameHeight);

                            Excel.Range cellDepartmentRange = excelSheet.Cells[i + 8, 3]; //Choose cell to Insert SignDepartment Image
                            float picsDepartmentWidth = 65;
                            float picsDepartmentHeight = 26;
                            double picsDepartmentLeft = cellDepartmentRange.Left + picsDepartmentWidth / 3;
                            double picsDepartmentTop = cellDepartmentRange.Top + 1;

                            excelSheet.Shapes.AddPicture(signDepartmentPath, LinkToFile, SaveWithDocument, (float)picsDepartmentLeft, (float)picsDepartmentTop, picsDepartmentWidth, picsDepartmentHeight);
                        }
                        catch
                        {
                            MessageBox.Show($"未找到{((AttendeesInfo)LvePeopleInfo.SelectedItems[i]).AttendeesName}的相应图片文件，请核对后添加!");
                        }
                    }
                    else
                    {
                        try
                        {
                            string signName = ((AttendeesInfo)LvePeopleInfo.SelectedItems[i]).AttendeesName;
                            string signNamePath = AppDomain.CurrentDomain.BaseDirectory + $@"..\\..\\DataResources\\SignImages\\{signName}姓名.svg"; ; // SignName Image Path

                            string signDepartmentPath = AppDomain.CurrentDomain.BaseDirectory + $@"..\\..\\DataResources\\SignImages\\{signName}部门.svg"; // SignDepartment Image Path

                            var LinkToFile = Microsoft.Office.Core.MsoTriState.msoFalse;
                            var SaveWithDocument = Microsoft.Office.Core.MsoTriState.msoTrue;
                            Excel.Range cellNameRange = excelSheet.Cells[i - 5, 5]; //Choose cell to Insert SignName Image
                            Pictures pics = (Pictures)excelSheet.Pictures(omissing);
                            float picsSignNameWidth = 40;
                            float picsSignNameHeight = 22.857f;
                            double picsSignNameLeft = cellNameRange.Left + picsSignNameWidth / 2;
                            double picsSignNameTop = cellNameRange.Top;

                            excelSheet.Shapes.AddPicture(signNamePath, LinkToFile, SaveWithDocument, (float)picsSignNameLeft, (float)picsSignNameTop, picsSignNameWidth, picsSignNameHeight);

                            Excel.Range cellDepartmentRange = excelSheet.Cells[i - 5, 6]; //Choose cell to Insert SignDepartment Image
                            float picsDepartmentWidth = 60;
                            float picsDepartmentHeight = 24;
                            double picsDepartmentLeft = cellDepartmentRange.Left + picsDepartmentWidth / 2;
                            double picsDepartmentTop = cellDepartmentRange.Top;

                            excelSheet.Shapes.AddPicture(signDepartmentPath, LinkToFile, SaveWithDocument, (float)picsDepartmentLeft, (float)picsDepartmentTop, picsDepartmentWidth, picsDepartmentHeight);
                        }
                        catch
                        {
                            MessageBox.Show($"未找到{((AttendeesInfo)LvePeopleInfo.SelectedItems[i]).AttendeesName}的相应图片文件，请核对后添加!");
                        }
                    }
                }
            }
        }

        private void CbxProfession_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ListView_Update();
        }

        private void CbxDepartment_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ListView_Update();
        }

        private void BtHelp_Click(object sender, RoutedEventArgs e)
        {
            string helpInfo =
                "1、本程序由刘勇开发完成(qq:2470687597)，仅用于辅助部门培训签到及年终培训考核！\n" +
                "2、本程序可增删部门人员文本信息，但对应人员姓名/部门/电话签名图片(.svg格式)需手动添加到安装程序相应文件中内；\n" +
                "3、本程序人员姓名签名图片大小：350*200像素，部门签名图片大小：500*200像素，电话签名图片大小：500*200像素；\n" +
                "4、所有图片均需去除背景后方可使用；\n" +
                "5、新增人员信息后需对当前窗口人员信息表单操作后人员信息表单内容才会更新！\n" +
                "6、禁止非本单位个人或团体下载使用本程序，本单位个人或团体不得用于从事任何商业、非法或侵权活动，由此对他人造成的权利侵害由用户自行承担；\n" +
                "鉴于用户设备软、硬件环境差异及本软件自身缺陷，作者不为使用本软件造成的损失承担任何责任；用户对本软件及共相关服务的使用将视为接受以上全部条款。";
            MessageBox.Show(helpInfo);
        }

        private void BtGenerate_Click(object sender, RoutedEventArgs e)
        {
            if (LvePeopleInfo.SelectedItems.Count == 0)
            {
                MessageBox.Show("请选择人员...");
            }
            else
            {
                WriteToExcel();
            }
        }

        private void BtAddItem_Click(object sender, RoutedEventArgs e)
        {
            AddItem pftraditemWindow = new AddItem();
            pftraditemWindow.Show();
        }

        private void BtReduceItem_Click(object sender, RoutedEventArgs e)
        {
            List<string> newLines = new List<string>();
            if (LvePeopleInfo.SelectedItems.Count > 0)
            {
                // Add Details to ListView
                StreamReader sr = new StreamReader(AppDomain.CurrentDomain.BaseDirectory + @"..\\..\\DataResources\\PersonnelList.db");
                string[] lines = sr.ReadToEnd().Split('\n');
                sr.Close();
                File.Delete(AppDomain.CurrentDomain.BaseDirectory + @"..\\..\\DataResources\\PersonnelList.db");

                foreach (string line in lines)
                {
                    for (int i = 0; i < LvePeopleInfo.SelectedItems.Count; i++)
                    {
                        string signName = ((AttendeesInfo)LvePeopleInfo.SelectedItems[i]).AttendeesName;
                        if (!line.Contains(signName))
                        {
                            newLines.Add(line.Trim('\n', ' ', '\r'));
                        }
                    }
                }
                newLines.RemoveAll(string.IsNullOrEmpty);
                List<string> distinctnewLines = newLines.Distinct().ToList();
                foreach (string newline in distinctnewLines)
                {
                    using (StreamWriter sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + @"..\\..\\DataResources\\PersonnelList.db", true))
                    {
                        sw.Write(newline + '\n');
                    }
                }
            }
            else
            {
                MessageBox.Show("未选中任何项...");
            }

            // Update ListView Contents
            ListView_Update();
        }
    }
}
