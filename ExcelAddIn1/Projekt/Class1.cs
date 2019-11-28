using System;

namespace ExcelAddIn.Model
{
    public class Projekt
    {
        public string Nazwa { get; set; }
        public double Ilosc { get; set; }
        public Projekt(string s, double d)
        {
            Nazwa = s;
            Ilosc = d;
        }
    }
}
