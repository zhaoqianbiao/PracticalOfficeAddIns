# PracticalOfficeAddIns
These are some practical office addins, aiming to bring much convenience when dealing with documents.​  
为了高效地编辑文档，开发了一些Office加载项

# Update Logs：
## 12/24/2024
The first addin, used to open the folder of the documents.(*docx / *doc)  
梦开始的地方：开发了一个打开文件保存目录的工具

Since the new version of office supports auto-saving to OneDrive, the program can automatically recognize the mode.  
由于新版的Office会自动保存文件，获取到的文件路径是在云上的（包含d.docs.live.net），如果你开启了OneDrive自动保存，需要根据个人的情况修改路径相关的部分代码

开发的时候考虑到python速度的问题以及无法定位只读文件，所以使用和Windows兼容性更好的C#

## 12/29/2024
全新开发！！
Excel工作表查找：根据工作表名称查找并跳转到对应的工作表
