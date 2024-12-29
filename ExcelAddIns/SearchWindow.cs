using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIns
{
    public partial class SearchWindow : Form
    {
        public SearchWindow()
        {
            InitializeComponent();

            // 设置窗口图标
            this.Icon = Properties.Resources.search;
            this.pictureBox1.Image = Properties.Resources.smallPig;

            // 设置窗体启动位置为手动
            this.StartPosition = FormStartPosition.Manual;
            // 获取屏幕分辨率
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            int screenHeight = Screen.PrimaryScreen.Bounds.Height;

            // 计算窗口位置，例如将窗口放置在屏幕中央
            int windowWidth = this.Width;
            int windowHeight = this.Height;
            int x = 80; // 左边缘
            int y = screenHeight - windowHeight - 150; // 下边缘

            // 设置窗体的位置
            this.Location = new Point(x, y);
        }

        private void SearchWindow_Load(object sender, EventArgs e)
        {
            // 列出所有的工作表
            foreach (Excel.Worksheet sheet in Globals.ThisAddIn.Application.ActiveWorkbook.Sheets)
            {
                listBox1.Items.Add(sheet.Name);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            /*
            // 检查是否处于输入法模式
            if (textBox1.ImeMode == ImeMode.On)
            {
                // 输入法临时内容
                MessageBox.Show("输入法临时内容");
            }
            else
            {
                // 真实的用户输入内容
                MessageBox.Show("真实的用户输入内容");
            }
            */

            // 获取输入的内容
            string searchText = textBox1.Text.ToLower();

            // 清空列表框
            listBox1.Items.Clear();

            // 重新列出匹配的工作表
            foreach (Excel.Worksheet sheet in Globals.ThisAddIn.Application.ActiveWorkbook.Sheets)
            {
                if (sheet.Name.ToLower().Contains(searchText))
                {
                    listBox1.Items.Add(sheet.Name);
                }
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 获取选中的工作表名称
            string selectedSheetName = listBox1.SelectedItem.ToString();

            // 查找并激活相应的工作表
            foreach (Excel.Worksheet sheet in Globals.ThisAddIn.Application.ActiveWorkbook.Sheets)
            {
                if (sheet.Name == selectedSheetName)
                {
                    sheet.Activate();
                    break;
                }
            }

            // 关闭搜索窗口
            this.Close();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
