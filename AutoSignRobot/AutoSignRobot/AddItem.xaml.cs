using System;
using System.Collections.Generic;
using System.IO;
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

namespace AutoSignRobot
{
    /// <summary>
    /// AddItem.xaml 的交互逻辑
    /// </summary>
    public partial class AddItem : Window
    {
        public AddItem()
        {
            InitializeComponent();
        }

        private void TbOK_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                List<string> newLines = new List<string>();
                // Read lines from Database
                StreamReader sr = new StreamReader(AppDomain.CurrentDomain.BaseDirectory + @"..\\..\\DataResources\\PersonnelList.db");
                string[] lines = sr.ReadToEnd().Split('\n');
                sr.Close();

                foreach (string line in lines)
                {
                    if (!line.Contains(TbxName.Text))
                    {
                        newLines.Add('\n' + TbxName.Text.Trim() + '\t' + TbxDepartment.Text.Trim() + '\t' + TbxProfession.Text.Trim() + '\t' + TbxTelePhone.Text.Trim() + '\t' + TbxPltafftion.Text.Trim());
                    }
                }

                using (StreamWriter sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + @"..\\..\\DataResources\\PersonnelList.db", true))
                {
                    foreach (string newline in newLines.Distinct().ToList())
                    {
                        sw.Write(newline);
                    }
                }

                // Update ListView Contents
                Close();
            }
            catch
            {
                MessageBox.Show($"输入为空，请核对后重新输入...");
            }
        }

        private void TbQuit_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
