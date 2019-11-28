using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelAddIn.Model.Interfaces
{
    public interface IDisplay
    {
        string Title { get; set; }
        double Hours { get; set; }
    }
}
