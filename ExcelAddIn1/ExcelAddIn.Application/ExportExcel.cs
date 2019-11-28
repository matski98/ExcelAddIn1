using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelAddIn.Model;
using ExcelAddIn.Model.Interfaces;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn.Application
{
    public class ExportExcel: IExportExcel
    {
        Microsoft.Office.Interop.Excel.Application _application { get; set; }
        public ExportExcel(Microsoft.Office.Interop.Excel.Application application)
        {
            _application = application;
        }
        public void Export(IExcelAddInModel model)
        {
            Worksheet worksheet = _application.Worksheets.Add();
            try
            {
                worksheet.Name = model.Project_Name;

            }
            catch
            {

            }
            worksheet.Cells[1, 1].Value = "Osoby\\Projekty";
            List<string> projects = new List<string>();
            foreach(IPerson person in model.Persons)
            {
                foreach(IProject project in person.Projects)
                {
                    if (!projects.Contains(project.Name))
                    {
                        projects.Add(project.Name);
                    }
                }
            }
            projects.Sort();
            for(int i = 2; i <= projects.Count+1; ++i)
            {
                worksheet.Cells[1, i].Value = projects[i-2];
            }
            int row = 2;
            foreach(IPerson person in model.Persons)
            {
                worksheet.Cells[row, 1].Value = person.Name +' '+ person.Surname;
                foreach(IProject project in person.Projects)
                {
                    for(int i = 0; i < projects.Count; ++i)
                    {
                        if (project.Name == projects[i])
                        {
                            worksheet.Cells[row, i + 2].Value = project.Number;
                        }
                    }
                }
                ++row;
            }
            worksheet.Columns.AutoFit();
            Range c1 = worksheet.Cells[1, 1];
            Range c2 = worksheet.Cells[row-1, worksheet.UsedRange.Columns.Count];
            Range tRange = (Range)worksheet.get_Range(c1, c2);

            worksheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, tRange, Type.Missing, XlYesNoGuess.xlYes, Type.Missing).Name = "TestTable";
            worksheet.ListObjects["TestTable"].TableStyle = "TableStyleMedium2";
        }
    }
}
