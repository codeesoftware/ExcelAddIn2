using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn2
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Log_Click(object sender, RibbonControlEventArgs e)
        {
            AccessActions accessActions = new AccessActions(@"D:\KPMG\probakörnyezet\Log.accdb");
            accessActions.ReadDataFromFile();
        }

        private void Download_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelOperations ExcelActions = new ExcelOperations(@"D:\KPMG\SvcUtil.exe");

            ExcelActions.Save();

        }
    }
}
