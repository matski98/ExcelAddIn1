using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelAddIn.Model.Interfaces
{
    public interface IExcelAddInModel
    {
        string Project_Name { get; set; }
        List<IPerson> Persons { get; set; }
        
    }
}
