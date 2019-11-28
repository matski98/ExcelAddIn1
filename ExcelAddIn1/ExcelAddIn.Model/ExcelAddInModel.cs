using ExcelAddIn.Model.Interfaces;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelAddIn.Model
{
    public class ExcelAddInModel : IExcelAddInModel
    {
        public List<IPerson> Persons { get; set; }
        public string Project_Name { get; set; }
        public ExcelAddInModel()
        {
            Persons = new List<IPerson>();
        }
    }
}