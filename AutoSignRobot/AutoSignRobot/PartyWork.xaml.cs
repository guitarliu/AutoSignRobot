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
using Word = Microsoft.Office.Interop.Word;
using Table = Microsoft.Office.Interop.Word.Table;

namespace AutoSignRobot
{
    /// <summary>
    /// PartyWork.xaml 的交互逻辑
    /// </summary>
    public partial class PartyWork : UserControl
    {
        public PartyWork()
        {
            InitializeComponent();

            // Add Details to ListView
            StreamReader sr = new StreamReader(AppDomain.CurrentDomain.BaseDirectory + @"..\\..\\DataResources\\PersonnelList.db");
            string[] lines = sr.ReadToEnd().Split('\n');
            sr.Close();

            List<string> cmbxProfession = new List<string> { };
            List<string> cmbxPltAfftion = new List<string> { };

            foreach (string line in lines)
            {
                if (line.Trim() == string.Empty)
                {
                    continue;
                }

                cmbxPltAfftion.Add(line.Split('\t')[4]);
                cmbxProfession.Add(line.Split('\t')[2]);

                AttendeesInfo attdInfo = new AttendeesInfo
                {
                    AttendeesName = line.Split('\t')[0],
                    AttendeesDepartment = line.Split('\t')[1],
                    AttendeesProfession = line.Split('\t')[2],
                    AttendeesPltaffiliation = line.Split('\t')[4].Trim('\n', '\r')
                };

                LvePeopleInfo.Items.Add(attdInfo);
            }

            string[] newCbxPltAfftion = cmbxPltAfftion.Distinct().ToList().ToArray();
            string[] newCbxProfession = cmbxProfession.Distinct().ToList().ToArray();
            CbxProfession.ItemsSource = newCbxProfession;
            CbxPltAfftion.ItemsSource = newCbxPltAfftion;
        }
        public class AttendeesInfo
        {
            public string AttendeesName { get; set; }
            public string AttendeesDepartment { get; set; }
            public string AttendeesProfession { get; set; }
            public string AttendeesPltaffiliation { get; set; }
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
                else if ((CbxProfession.SelectedIndex == -1) & (CbxPltAfftion.SelectedIndex != -1))
                {
                    if (CbxPltAfftion.SelectedItem.ToString() == line.Split('\t')[4])
                    {
                        AttendeesInfo attdInfo = new AttendeesInfo
                        {
                            AttendeesName = line.Split('\t')[0],
                            AttendeesDepartment = line.Split('\t')[1],
                            AttendeesProfession = line.Split('\t')[2],
                            AttendeesPltaffiliation = line.Split('\t')[4].Trim('\n', '\r')
                        };

                        LvePeopleInfo.Items.Add(attdInfo);
                    }
                }
                else if ((CbxProfession.SelectedIndex != -1) & (CbxPltAfftion.SelectedIndex == -1))
                {
                    if (CbxProfession.SelectedItem.ToString() == line.Split('\t')[2])
                    {
                        AttendeesInfo attdInfo = new AttendeesInfo
                        {
                            AttendeesName = line.Split('\t')[0],
                            AttendeesDepartment = line.Split('\t')[1],
                            AttendeesProfession = line.Split('\t')[2],
                            AttendeesPltaffiliation = line.Split('\t')[4].Trim('\n', '\r')
                        };

                        LvePeopleInfo.Items.Add(attdInfo);
                    }
                }
                else if ((CbxProfession.SelectedIndex != -1) & (CbxPltAfftion.SelectedIndex != -1))
                {
                    if ((CbxProfession.SelectedItem.ToString() == line.Split('\t')[2]) & (CbxPltAfftion.SelectedItem.ToString() == line.Split('\t')[4]))
                    {
                        AttendeesInfo attdInfo = new AttendeesInfo
                        {
                            AttendeesName = line.Split('\t')[0],
                            AttendeesDepartment = line.Split('\t')[1],
                            AttendeesProfession = line.Split('\t')[2],
                            AttendeesPltaffiliation = line.Split('\t')[4].Trim('\n', '\r')
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
                        AttendeesPltaffiliation = line.Split('\t')[4].Trim('\n', '\r')
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

            if (!DpStartDate.SelectedDate.HasValue)
            {
                MessageBox.Show("请选择日期及时间！", "错误");
            }
            else
            {
                DateResult = $"{DpStartDate.SelectedDate.Value.ToString("D")}";
                var wordApp = new Word.Application();
                var wordDoc = new Document();
                wordApp.Visible = true;
                wordDoc = wordApp.Documents.Add(ref omissing, ref omissing, ref omissing,
                ref omissing);

                // Set Document Title Font Style
                wordApp.Selection.EndKey(ref unite, ref omissing);//Move Mouse to the end of line
                wordDoc.Paragraphs.Last.Range.Font.Name = "方正小标宋简体";
                wordDoc.Paragraphs.Last.Range.Font.Size = 16;
                wordDoc.Paragraphs.Last.Range.Font.Bold = 0;
                wordDoc.Paragraphs.Last.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                wordDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                wordDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacing = 30;

                wordDoc.Paragraphs.Last.Range.Text = TbxTitle.Text + "\n";

                // Set Document Content Font Style
                wordApp.Selection.EndKey(ref unite, ref omissing);//Move Mouse to the end of line
                wordDoc.Paragraphs.Last.Range.Font.Name = "华文仿宋";
                wordDoc.Paragraphs.Last.Range.Font.Size = 12;
                wordDoc.Paragraphs.Last.Range.Font.Bold = 0;
                wordDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                wordDoc.Paragraphs.Last.Range.ParagraphFormat.LineSpacing = 21;
                wordDoc.Paragraphs.Last.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                wordDoc.Paragraphs.Last.Range.ParagraphFormat.LineUnitBefore = (float)0.5;

                wordDoc.Paragraphs.Last.Range.Text = $"时间：{DateResult}\n";
                wordDoc.Paragraphs.Last.Range.Text = $"地点：{TbxLocation.Text}\n";
                wordDoc.Paragraphs.Last.Range.Text = $"议程：{TbxProceedings.Text}\n";

                // Define Table Row and Column Number
                int tableRow = 12;
                int tableColumn = 8;
                // Add Table
                wordApp.Selection.EndKey(ref unite, ref omissing);//Move Mouse to the end of line
                Table table = wordDoc.Tables.Add(wordApp.Selection.Range,
                    tableRow, tableColumn, ref omissing, ref omissing);
                table.Borders.Enable = 1;

                // Set Table Title List
                List<string> tableTitle = new List<string>
                {
                "编号",
                "姓  名",
                "签  到",
                "备 注"
                };

                // Set Table Columns and Rows Width and Height
                Dictionary<int, float> columnWidthDict = new Dictionary<int, float>
                {
                    {1, 8.0f},
                    {2, 13.9f},
                    {3, 13.9f},
                    {4, 13.9f},
                    {5, 8.0f},
                    {6, 13.9f},
                    {7, 13.9f},
                    {8, 13.9f}
                };

                for (int i = 1; i < tableColumn - 3; i++)
                {
                    wordDoc.Tables[1].Cell(1, i).Range.Font.Name = "方正风雅楷宋简体";
                    wordDoc.Tables[1].Cell(1, i).Range.Font.Size = 12;
                    wordDoc.Tables[1].Cell(1, i).Range.Text = tableTitle[i - 1];
                    wordDoc.Tables[1].Cell(1, i).Range.Font.Bold = 10;
                }
                for (int i = 5; i < tableColumn + 1; i++)
                {
                    wordDoc.Tables[1].Cell(1, i).Range.Font.Name = "方正风雅楷宋简体";
                    wordDoc.Tables[1].Cell(1, i).Range.Font.Size = 12;
                    wordDoc.Tables[1].Cell(1, i).Range.Text = tableTitle[i - 5];
                    wordDoc.Tables[1].Cell(1, i).Range.Font.Bold = 10;
                }
                for (int i = 2; i < tableRow + 1; i++)
                {
                    wordDoc.Tables[1].Cell(i, 1).Range.Font.Name = "方正悠宋 简 507R";
                    wordDoc.Tables[1].Cell(i, 1).Range.Font.Size = 12;
                    wordDoc.Tables[1].Cell(i, 5).Range.Font.Name = "方正悠宋 简 507R";
                    wordDoc.Tables[1].Cell(i, 5).Range.Font.Size = 12;

                    wordDoc.Tables[1].Cell(i, 2).Range.Font.Name = "宋体";
                    wordDoc.Tables[1].Cell(i, 2).Range.Font.Size = 12;
                    wordDoc.Tables[1].Cell(i, 6).Range.Font.Name = "宋体";
                    wordDoc.Tables[1].Cell(i, 6).Range.Font.Size = 12;

                    wordDoc.Tables[1].Cell(i, 1).Range.Text = $"{i - 1}";
                    wordDoc.Tables[1].Cell(i, 5).Range.Text = $"{i + 10}";
                }

                // table content keep center
                table.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                table.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                // Set table Title linespacing rule
                table.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                table.Range.ParagraphFormat.LineUnitAfter = 0;
                table.Range.ParagraphFormat.LineUnitBefore = 0;
                wordDoc.Tables[1].Rows.Alignment = WdRowAlignment.wdAlignRowCenter;

                // Set table Columns' Width
                foreach (KeyValuePair<int, float> kvp in columnWidthDict)
                {
                    table.Columns[kvp.Key].PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
                    table.Columns[kvp.Key].PreferredWidth = kvp.Value;
                }

                // Set table Rows' Height
                table.Rows.HeightRule = WdRowHeightRule.wdRowHeightAtLeast;
                for (int i = 1; i < tableRow + 1; i++)
                {
                    table.Rows[i].Height = wordApp.CentimetersToPoints(1.52f);
                    table.Rows[i].Range.ParagraphFormat.LineUnitBefore = 0;
                    table.Rows[i].Range.ParagraphFormat.LineUnitAfter = 0;
                    table.Rows[i].Range.ParagraphFormat.SpaceBefore = 0;
                }


                /// Set whole table width
                /// Only Can do this when all operations of changing table columns width ended
                wordDoc.Tables[1].PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
                wordDoc.Tables[1].PreferredWidth = 107f;

                // Write Pictures to Table
                // According to ListView Selected Items
                for (int i = 0; i < LvePeopleInfo.SelectedItems.Count; i++)
                {
                    if (i <= 10)
                    {
                        try
                        {
                            string signName = ((AttendeesInfo)LvePeopleInfo.SelectedItems[i]).AttendeesName;
                            string signPltaffiliation = ((AttendeesInfo)LvePeopleInfo.SelectedItems[i]).AttendeesPltaffiliation.Trim('\n', '\r', ' ');
                            string signNamePath = AppDomain.CurrentDomain.BaseDirectory + $@"..\\..\\DataResources\\SignImages\\{signName}姓名.svg"; ; // SignName Image Path

                            object LinkToFile = false;
                            object SaveWithDocument = true;

                            if (signName.Length < 3) { wordDoc.Tables[1].Cell(i + 2, 2).Range.Text = $"{signName[0]}" + "  " + $"{signName[1]}"; }
                            else { wordDoc.Tables[1].Cell(i + 2, 2).Range.Text = $"{signName}"; }

                            if ((signPltaffiliation != "党员") & (!signPltaffiliation.Contains("群众"))) { wordDoc.Tables[1].Cell(i + 2, 4).Range.Text = $"{signPltaffiliation}"; }

                            object cellNameRange = table.Cell(i + 2, 3).Range; //Choose cell to Insert SignName Image
                            wordDoc.Application.ActiveDocument.InlineShapes.AddPicture(signNamePath, ref LinkToFile, ref SaveWithDocument, ref cellNameRange);
                            wordDoc.Application.ActiveDocument.InlineShapes[i + 1].Width = 40; //Set Picture Width
                            wordDoc.Application.ActiveDocument.InlineShapes[i + 1].Height = 22.875f; //Set Picture Height
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
                            string signPltaffiliation = ((AttendeesInfo)LvePeopleInfo.SelectedItems[i]).AttendeesPltaffiliation.Trim('\n', '\r', ' ');
                            string signNamePath = AppDomain.CurrentDomain.BaseDirectory + $@"..\\..\\DataResources\\SignImages\\{signName}姓名.svg"; ; // SignName Image Path

                            object LinkToFile = false;
                            object SaveWithDocument = true;

                            if (signName.Length < 3) { wordDoc.Tables[1].Cell(i - 9, 6).Range.Text = $"{signName[0]}" + "  " + $"{signName[1]}"; }
                            else { wordDoc.Tables[1].Cell(i - 9, 6).Range.Text = $"{signName}"; }

                            if ((signPltaffiliation != "党员") & (!signPltaffiliation.Contains("群众"))) { wordDoc.Tables[1].Cell(i - 9, 8).Range.Text = $"{signPltaffiliation}"; }

                            object cellNameRange = table.Cell(i - 9, 7).Range; //Choose cell to Insert SignName Image
                            wordDoc.Application.ActiveDocument.InlineShapes.AddPicture(signNamePath, ref LinkToFile, ref SaveWithDocument, ref cellNameRange);
                            wordDoc.Application.ActiveDocument.InlineShapes[i + 1].Width = 50; //Set Picture Width
                            wordDoc.Application.ActiveDocument.InlineShapes[i + 1].Height = 28.571f; //Set Picture Height
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
        private void CbxPltAfftion_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ListView_Update();
        }
        private void BtGenerate_Click(object sender, RoutedEventArgs e)
        {
            if (LvePeopleInfo.SelectedItems.Count == 0)
            {
                MessageBox.Show("请选择党组织相关人员...");
            }
            else
            {
                WriteToWord();
            }
        }

        private void BtHelp_Click(object sender, RoutedEventArgs e)
        {
            string helpInfo =
                "1、本程序由刘勇开发完成（qq:2470687597），仅用于辅助部门党小组会议签到！\n" +
                "2、本程序可增删部门人员文本信息，但对应人员姓名/部门/电话签名图片(.svg格式)需手动添加到安装程序相应文件中内；\n" +
                "3、本程序人员姓名签名图片大小：350*200像素，部门签名图片大小：500*200像素，电话签名图片大小：500*200像素；\n" +
                "4、所有图片均需去除背景后方可使用；\n" +
                "5、新增人员信息后需对当前窗口人员信息表单操作后人员信息表单内容才会更新！\n" +
                "6、禁止非本单位个人或团体下载使用本程序，本单位个人或团体不得用于从事任何商业、非法或侵权活动，由此对他人造成的权利侵害由用户自行承担；\n" +
                "鉴于用户设备软、硬件环境差异及本软件自身缺陷，作者不为使用本软件造成的损失承担任何责任；用户对本软件及共相关服务的使用将视为接受以上全部条款。";
            MessageBox.Show(helpInfo);
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

        private void BtAddItem_Click(object sender, RoutedEventArgs e)
        {
            AddItem pftraditemWindow = new AddItem();
            pftraditemWindow.Show();
        }

    }
}
