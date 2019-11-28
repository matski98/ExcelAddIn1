using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelAddIn.Model.Interfaces
{
    public interface IPerson
    {
            string Name { get; set; }
            string Surname{ get; set; }
            List<IProject> Projects { get; set; }
    }
}
