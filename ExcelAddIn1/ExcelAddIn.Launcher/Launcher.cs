using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using ExcelAddIn.Application;
using ExcelAddIn.Framework;
using ExcelAddIn.Model;
using ExcelAddIn.Model.Interfaces;
using ExcelAddIn.GUI.Interfaces;
using ExcelAddIn.GUI;
using ExcelAddIn.Presenter;

namespace ExcelAddIn.Launcher
{
    public class Launcher
    {
        Worksheet CurrentSheet { get; set; }
        Microsoft.Office.Interop.Excel.Application _application { get; set; }
        IExcelAddInView view;
        IExcelAddInModel model;
        public Launcher(Worksheet currentSheet, Microsoft.Office.Interop.Excel.Application app)
        {
            _application = app;
            StartServices();
            CurrentSheet = currentSheet;


        }

        public void Run()
        {
            IReaderService _readingservice = ServiceProvider.GetService<IReaderService>();
            
            model = _readingservice.Read(CurrentSheet);
            view = new ExcellAddInView();

            ExcelAddInPresenter presenter = new ExcelAddInPresenter(view, model);

        }

        public void StartServices()
        {
            ServiceProvider.RegisterService(typeof(IReaderService), new ReaderService());
            ServiceProvider.RegisterService(typeof(IExportExcel), new ExportExcel(_application));
        }

        public void Export(Worksheet worksheet)
        {
            IExportExcel exportExcel = ServiceProvider.GetService<IExportExcel>();
        }
    }
}
