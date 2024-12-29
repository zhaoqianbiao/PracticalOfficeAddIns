using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelAddIns
{
    public partial class QbTools
    {
        private void QbTools_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Search搜索工作表_Click(object sender, RibbonControlEventArgs e)
        {
            new SearchWindow().Show();
        }
    }
}
