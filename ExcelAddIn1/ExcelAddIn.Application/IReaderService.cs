using ExcelAddIn.Framework;
using ExcelAddIn.Model;
using ExcelAddIn.Model.Interfaces;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn.Application
{
    public interface IReaderService : IService
    {
        IExcelAddInModel Read(Worksheet currentSheet);
    }
}