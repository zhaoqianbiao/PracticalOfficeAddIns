using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WordAddIns
{
    public partial class Zone1
    {
        private void Zone1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        public static string ConvertOneDriveUrlToLocalPath(string url)
        {
            // 替换为你的 OneDrive 本地文件夹路径
            string onedriveLocalPath = Environment.ExpandEnvironmentVariables(@"C:\Users\qianb\OneDrive");
            /*MessageBox.Show($"OneDrive local path: {onedriveLocalPath}");*/

            try
            {
                // 提取 URL 中的相对路径
                Uri uri = new Uri(url);
                string relativePath = uri.AbsolutePath.Substring(uri.AbsolutePath.IndexOf("/d.docs.live.net/") + "/d.docs.live.net/".Length);
                relativePath = relativePath.Substring(relativePath.IndexOf("/") + 1);

                // 解码 URL 中的特殊字符
                relativePath = Uri.UnescapeDataString(relativePath);

                // 构建本地路径
                string localPath = Path.Combine(onedriveLocalPath, relativePath.Replace("/", Path.DirectorySeparatorChar.ToString()));
                localPath = Path.GetFullPath(localPath);  // 规范化路径以使用正确的分隔符
                return localPath;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: Unable to parse OneDrive URL. {ex.Message}");
                return null;
            }
        }

        private string GetActiveDocumentPath()
        {
            // 获取当前活动文档的完整路径
            var activeDocument = Globals.ThisAddIn.Application.ActiveDocument;
            if (activeDocument != null)
            {
                return activeDocument.FullName;
            }
            return null;
        }

        private void OpenInFolder_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前活动文档的路径
            string activeDocumentPath = GetActiveDocumentPath();

            if (activeDocumentPath.Contains("d.docs.live.net"))
            {
                activeDocumentPath = ConvertOneDriveUrlToLocalPath(activeDocumentPath);
            }


            if (!string.IsNullOrEmpty(activeDocumentPath))
            {
                // 获取文件所在的目录
                string directoryPath = Path.GetDirectoryName(activeDocumentPath);

                //MessageBox.Show("OneDrive Path: " + directoryPath);

                if (Directory.Exists(directoryPath))
                {
                    // 打开文件所在的目录
                    //Process.Start("explorer.exe", directoryPath);
                    Process.Start("explorer.exe", $"/select,\"{activeDocumentPath}\"");
                }
                else
                {
                    MessageBox.Show("目录不存在。");
                }
            }
            else
            {
                MessageBox.Show("未找到活动文档。");
            }
        }
    }
}
