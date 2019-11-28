using ExcelAddIn.Model;
using ExcelAddIn.Model.Interfaces;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAddIn.Application
{
    public class ReaderService : IReaderService
    {
        public IExcelAddInModel Read(Worksheet currentSheet)
        {
            Range usedR = currentSheet.UsedRange;
            IExcelAddInModel model = new ExcelAddInModel();
            string project_name = "";
            model.Project_Name = usedR.Cells[7, 2].Value.ToString().Split(',')[1].Remove(0,1);
            for (int row = 1; row <= usedR.Rows.Count; row++)
            {
                if (usedR.Cells[row, 2].Value != null)
                {
                    var celldata1 = usedR.Cells[row, 2].Value;
                    celldata1 = celldata1.ToString();
                    string[] celldata1_tab = celldata1.Split(' ');
                    if (celldata1_tab[0].Length >= 18)
                    {
                        if (celldata1_tab[0][13] == '-' && celldata1_tab[0][17] == '-')
                        {
                            project_name = celldata1_tab[0];
                        }
                    }
                }
                var celldata2 = usedR.Cells[row, 1].Value;
                if (celldata2 != null)
                {
                    celldata2 = celldata2.ToString();
                    string[] celldata2_tab = celldata2.Split(' ');
                    if (celldata2_tab[0] == "Summe" && Char.IsDigit(celldata2_tab[1][0]))
                    {
                        bool not_appear_before = true;
                        if (model.Persons.Count > 0)
                        {
                            foreach (IPerson person in model.Persons)
                            {
                                if (person.Surname == celldata2_tab[3].Remove(celldata2_tab[3].Length - 1, 1) && person.Name == celldata2_tab[4])
                                {
                                    not_appear_before = false;
                                    double liczba_h = usedR.Cells[row, 6].Value;
                                    IProject p = new Project(project_name, liczba_h);
                                    person.Projects.Add(p);
                                }
                            }
                        }
                        AddIfNotAppearBefore(usedR, model, project_name, row, celldata2_tab, not_appear_before);
                    }
                }
            }
            return model;
        }

        private static void AddIfNotAppearBefore(Range usedR, IExcelAddInModel model, string project_name, int row, string[] celldata_tab, bool not_appear_before)
        {
            if (not_appear_before)
            {
                double hours = usedR.Cells[row, 6].Value;
                IProject project = new Project(project_name, hours);
                IPerson person = new Person() { Surname = celldata_tab[3].Remove(celldata_tab[3].Length - 1, 1), Name = celldata_tab[4] };
                person.Projects.Add(project);
                model.Persons.Add(person);
            }
        }
    }
}