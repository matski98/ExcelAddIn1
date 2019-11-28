using ExcelAddIn.Model.Interfaces;
using System;
using System.Collections.Generic;


namespace ExcelAddIn.Model
{
    public class Display : IDisplay
    { 
        public string Title { get; set; }
        public double Hours { get; set; }

        public override string ToString()
        {
            return Title;
        }
    }
}