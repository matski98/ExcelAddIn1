using ExcelAddIn.Model.Interfaces;
using System;

namespace ExcelAddIn.Model
{
    public class Project : IProject
    {
        public string Name { get; set; }
        public double Number { get; set; }   
        public Project(string s, double d)
        {
            Name = s;
            Number = d;
        }
    }
}
