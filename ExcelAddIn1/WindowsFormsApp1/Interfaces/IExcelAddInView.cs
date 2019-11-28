using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelAddIn.Model.Interfaces;

namespace ExcelAddIn.GUI.Interfaces
{
    public interface IExcelAddInView
    {
        DialogResult Showdialog();

        void Close();

        event EventHandler Cancel;
        event EventHandler Export;
        event EventHandler SelectedIndex;
        event EventHandler ViewingMode;

        void SetupList(List<IDisplay> list);
        void SetupGrid(List<IDisplay> list);

        int ViewingModeValue{ get; }
        int SelectedIndexValue { get; }
    }
}
