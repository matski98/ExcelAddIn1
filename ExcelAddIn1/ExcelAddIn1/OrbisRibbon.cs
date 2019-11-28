using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections;
using ExcelAddIn.Model;
using ExcelAddIn.Launcher;

namespace ExcelAddIn.Excel
{

    public partial class OrbisRibbon
    {

        public void btnButton_Click(object sender, RibbonControlEventArgs e)
        {
            Application app = Globals.ThisAddIn.Application;
            Worksheet currentSheet = Globals.ThisAddIn.GetActiveWorkSheet();
            Launcher.Launcher launcher = new Launcher.Launcher(currentSheet, app);
            launcher.Run();
            //launcher.ExportExcel += ExportExcel;
            //Worksheet newWorksheet = Globals.ThisAddIn.AddWorkSheet();
            //launcher.Export(newWorksheet);
        }
        private void ExportExcel(object sender, EventArgs e)
        {
            Worksheet newWorksheet = Globals.ThisAddIn.AddWorkSheet();
        }
    }
}
