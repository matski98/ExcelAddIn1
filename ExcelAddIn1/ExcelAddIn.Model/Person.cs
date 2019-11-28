using ExcelAddIn.Model.Interfaces;
using System;
using System.Collections.Generic;


namespace ExcelAddIn.Model
{
    public class Person : IPerson
    {
        public string Name { get; set; }
        public string Surname { get; set; }
        public List<IProject> Projects { get; set; }
        public Person()
        {
            Projects = new List<IProject>();
            Surname = "";
            Name = "";
        }
    }
}