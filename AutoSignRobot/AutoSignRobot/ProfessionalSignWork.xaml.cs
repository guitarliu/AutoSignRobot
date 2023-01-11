using Microsoft.Office.Interop.Word;
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
using static AutoSignRobot.PartyWork;
using Word = Microsoft.Office.Interop.Word;
using Table = Microsoft.Office.Interop.Word.Table;

namespace AutoSignRobot
{
    /// <summary>
    /// ProfessionalSignWork.xaml 的交互逻辑
    /// </summary>
    public partial class ProfessionalSignWork : UserControl
    {
        public ProfessionalSignWork()
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

        public void WriteToWord()
        {
            // The Add method has four reference parameters, all of which are
            // optional. Visual C# allows you to omit arguments for them if
            // the default values are what you want.
            object omissing = Missing.Value;
            object unite = WdUnits.wdStory;
            string DateResult = "";

            if ((!DpStartDate.SelectedDate.HasValue) | (!DpEndDate.SelectedDate.HasValue) | (TbxStartTime.Text == "") | (TbxEndTime.Text == ""))
            {
                MessageBox.Show("请选择日期及时间！", "错误");
            }
            else
            {
                if (DpStartDate.SelectedDate.Value == DpEndDate.SelectedDate.Value)
                {
                    DateResult = $"{DpStartDate.SelectedDate.Value.ToString("D")}" +
                        $"{TbxStartTime.Text}—{TbxEndTime.Text}";
                }
                else
                {
                    DateResult = $"{DpStartDate.SelectedDate.Value.ToString("D")}{TbxStartTime.Text}—" +
                        $"{DpEndDate.SelectedDate.Value.ToString("D")}{TbxEndTime.Text}";
                }
                var wordApp = new Word.Application();
                var wordDoc = new Document();
                wordApp.Visible = true;
                wordDoc = wordApp.Documents.Add(ref omissing, ref omissing, ref omissing,
                ref omissing);

                // Set Document Title Font Style
                wordApp.Selection.EndKey(ref unite, ref omissing);//Move Mouse to the end of line
                wordDoc.Paragraphs.Last.Range.Font.Name = "宋体";
                wordDoc.Paragraphs.Last.Range.Font.Size = 14;
                wordDoc.Paragraphs.Last.Range.Font.Bold = 10;
                wordDoc.Paragraphs.Last.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                wordDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;
                wordDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacing = wordApp.LinesToPoints(4);

                wordDoc.Paragraphs.Last.Range.Text = TbxTitle.Text + "\n";

                // Set Document Content Font Style
                wordApp.Selection.EndKey(ref unite, ref omissing);//Move Mouse to the end of line
                wordDoc.Paragraphs.Last.Range.Font.Name = "宋体";
                wordDoc.Paragraphs.Last.Range.Font.Size = 12;
                wordDoc.Paragraphs.Last.Range.Font.Bold = 10;
                wordDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpace1pt5;
                wordDoc.Paragraphs.Last.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;

                wordDoc.Paragraphs.Last.Range.Text = $"起始时间：{DateResult}\t\t地点：{TbxLocation.Text}\n";
                wordDoc.Paragraphs.Last.Range.Text = $"主讲人：{TbxTalker.Text}\n";
                wordDoc.Paragraphs.Last.Range.Text = $"内容：{TbxContent.Text}\n";

                // Define Table Row and Column Number
                int tableRow = 18;
                int tableColumn = 8;
                // Add Table
                wordApp.Selection.EndKey(ref unite, ref omissing);//Move Mouse to the end of line
                Table table = wordDoc.Tables.Add(wordApp.Selection.Range,
                    tableRow, tableColumn, ref omissing, ref omissing);
                table.Borders.Enable = 1;

                // Set Table Title List
                List<string> tableTitle = new List<string>
                {
                "序号",
                "姓名",
                "部门",
                "电话"
                };

                // Set Table Columns and Rows Width and Height
                Dictionary<int, float> columnWidthDict = new Dictionary<int, float>
                {
                    {1, 0.95f},
                    {2, 1.6f},
                    {3, 2.58f},
                    {4, 2.45f},
                    {5, 0.99f},
                    {6, 1.53f},
                    {7, 2.58f},
                    {8, 2.60f}
                };

                for (int i = 1; i < tableColumn - 3; i++)
                {
                    wordDoc.Tables[1].Cell(1, i).Range.Font.Name = "宋体";
                    wordDoc.Tables[1].Cell(1, i).Range.Font.Size = 12;
                    wordDoc.Tables[1].Cell(1, i).Range.Text = tableTitle[i - 1];
                }
                for (int i = 5; i < tableColumn + 1; i++)
                {
                    wordDoc.Tables[1].Cell(1, i).Range.Font.Name = "宋体";
                    wordDoc.Tables[1].Cell(1, i).Range.Font.Size = 12;
                    wordDoc.Tables[1].Cell(1, i).Range.Text = tableTitle[i - 5];
                }
                for (int i = 2; i < tableRow + 1; i++)
                {
                    wordDoc.Tables[1].Cell(i, 1).Range.Font.Name = "Calibri";
                    wordDoc.Tables[1].Cell(i, 1).Range.Font.Size = 12;
                    wordDoc.Tables[1].Cell(i, 5).Range.Font.Name = "Calibri";
                    wordDoc.Tables[1].Cell(i, 5).Range.Font.Size = 12;

                    wordDoc.Tables[1].Cell(i, 1).Range.Text = $"{i - 1}.";
                    wordDoc.Tables[1].Cell(i, 5).Range.Text = $"{i + 16}.";
                }

                // table content keep center
                table.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                table.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                // Set table Title linespacing rule
                for (int i = 1; i < tableColumn + 1; i++)
                {
                    float lnePoints;
                    if ((i == 1) | (i == 5))
                    {
                        lnePoints = 1.5f;
                    }
                    else
                    {
                        lnePoints = 3f;
                    }
                    wordDoc.Tables[1].Cell(1, i).Range.ParagraphFormat.LineSpacing = wordApp.LinesToPoints(lnePoints);
                }

                // Set table Columns' Width
                foreach (KeyValuePair<int, float> kvp in columnWidthDict)
                {
                    table.Columns[kvp.Key].PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints;
                    table.Columns[kvp.Key].PreferredWidth = wordApp.CentimetersToPoints(kvp.Value);
                }

                // Set table Rows' Height
                table.Rows.HeightRule = WdRowHeightRule.wdRowHeightAtLeast;
                for (int i = 1; i < tableRow + 1; i++)
                {
                    if (i == 1)
                    {
                        table.Rows[i].Height = wordApp.CentimetersToPoints(1.58f);
                    }
                    else
                    {
                        table.Rows[i].Height = wordApp.CentimetersToPoints(0.98f);
                    }
                }

                // Write Pictures to Table
                // According to ListView Selected Items
                for (int i = 0; i < LvePeopleInfo.SelectedItems.Count; i++)
                {
                    if (i <= 16)
                    {
                        try
                        {
                            string signName = ((AttendeesInfo)LvePeopleInfo.SelectedItems[i]).AttendeesName;
                            string signNamePath = AppDomain.CurrentDomain.BaseDirectory + $@"..\\..\\DataResources\\SignImages\\{signName}姓名.svg"; ; // SignName Image Path

                            string signDepartmentPath = AppDomain.CurrentDomain.BaseDirectory + $@"..\\..\\DataResources\\SignImages\\{signName}部门.svg"; // SignDepartment Image Path

                            string signTelePhonePath = AppDomain.CurrentDomain.BaseDirectory + $@"..\\..\\DataResources\\SignImages\\{signName}电话.svg"; // SignTelePhone Image Path

                            object LinkToFile = false;
                            object SaveWithDocument = true;
                            object cellNameRange = table.Cell(i + 2, 2).Range; //Choose cell to Insert SignName Image
                            wordDoc.Application.ActiveDocument.InlineShapes.AddPicture(signNamePath, ref LinkToFile, ref SaveWithDocument, ref cellNameRange);
                            wordDoc.Application.ActiveDocument.InlineShapes[i + 1].Width = 40; //Set Picture Width
                            wordDoc.Application.ActiveDocument.InlineShapes[i + 1].Height = 22.875f; //Set Picture Height

                            object cellDepartmentRange = table.Cell(i + 2, 3).Range; //Choose cell to Insert SignDepartment Image
                            wordDoc.Application.ActiveDocument.InlineShapes.AddPicture(signDepartmentPath, ref LinkToFile, ref SaveWithDocument, ref cellDepartmentRange);
                            wordDoc.Application.ActiveDocument.InlineShapes[i + 2].Width = 55; //Set Picture Width
                            wordDoc.Application.ActiveDocument.InlineShapes[i + 2].Height = 22; //Set Picture Height

                            object cellTelePhoneRange = table.Cell(i + 2, 4).Range; //Choose cell to Insert SignTelePhone Image
                            wordDoc.Application.ActiveDocument.InlineShapes.AddPicture(signTelePhonePath, ref LinkToFile, ref SaveWithDocument, ref cellTelePhoneRange);
                            wordDoc.Application.ActiveDocument.InlineShapes[i + 3].Width = 55; //Set Picture Width
                            wordDoc.Application.ActiveDocument.InlineShapes[i + 3].Height = 22; //Set Picture Height
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

                            string signTelePhonePath = AppDomain.CurrentDomain.BaseDirectory + $@"..\\..\\DataResources\\SignImages\\{signName}电话.svg"; // SignTelePhone Image Path

                            object LinkToFile = false;
                            object SaveWithDocument = true;
                            object cellNameRange = table.Cell(i - 15, 6).Range; //Choose cell to Insert SignName Image
                            wordDoc.Application.ActiveDocument.InlineShapes.AddPicture(signNamePath, ref LinkToFile, ref SaveWithDocument, ref cellNameRange);
                            wordDoc.Application.ActiveDocument.InlineShapes[i + 1].Width = 40; //Set Picture Width
                            wordDoc.Application.ActiveDocument.InlineShapes[i + 1].Height = 22.875f; //Set Picture Height

                            object cellDepartmentRange = table.Cell(i - 15, 7).Range; //Choose cell to Insert SignDepartment Image
                            wordDoc.Application.ActiveDocument.InlineShapes.AddPicture(signDepartmentPath, ref LinkToFile, ref SaveWithDocument, ref cellDepartmentRange);
                            wordDoc.Application.ActiveDocument.InlineShapes[i + 2].Width = 55; //Set Picture Width
                            wordDoc.Application.ActiveDocument.InlineShapes[i + 2].Height = 22; //Set Picture Height

                            object cellTelePhoneRange = table.Cell(i - 15, 8).Range; //Choose cell to Insert SignTelePhone Image
                            wordDoc.Application.ActiveDocument.InlineShapes.AddPicture(signTelePhonePath, ref LinkToFile, ref SaveWithDocument, ref cellTelePhoneRange);
                            wordDoc.Application.ActiveDocument.InlineShapes[i + 3].Width = 55; //Set Picture Width
                            wordDoc.Application.ActiveDocument.InlineShapes[i + 3].Height = 22; //Set Picture Height
                        }
                        catch
                        {
                            MessageBox.Show($"未找到{((AttendeesInfo)LvePeopleInfo.SelectedItems[i]).AttendeesName}的相应图片文件，请核对后添加!");
                        }
                    }
                }
            }
        }

        private void CbxProfession_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            ListView_Update();
        }

        private void CbxDepartment_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            ListView_Update();
        }

        private void BtHelp_Click(object sender, RoutedEventArgs e)
        {
            string helpInfo =
                "1、本程序由刘勇开发完成（qq:2470687597），仅用于辅助部门培训签到及年终培训考核！\n" +
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
                WriteToWord();
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
