using ExcelAddIn.Framework;
using ExcelAddIn.Model;
using ExcelAddIn.Model.Interfaces;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn.Application
{
    public interface IExportExcel : IService
    {
        void Export(IExcelAddInModel model);
    }
}