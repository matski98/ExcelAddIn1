using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelAddIn.Model.Interfaces
{
    public interface IProject
    {
        string Name { get; set; }
        double Number { get; set; }
    }
}
