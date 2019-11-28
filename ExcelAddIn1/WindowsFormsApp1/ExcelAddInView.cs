using ExcelAddIn.GUI.Interfaces;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelAddIn.Model.Interfaces;
using ExcelAddIn.Model;

namespace ExcelAddIn.GUI
{
    public partial class ExcellAddInView : Form, IExcelAddInView
    {
        public event EventHandler Cancel;
        public event EventHandler Export;
        public event EventHandler ViewingMode;
        public event EventHandler SelectedIndex;
        public ExcellAddInView()
        {
            InitializeComponent();
            ViewingModeComboBox.SelectedItem = "Osoby";
        }

        public DialogResult Showdialog()
        {
            return ShowDialog();
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox1.SelectedItems.Count != 0 && SelectedIndex != null)
            {
                SelectedIndex(this, new EventArgs());
            }
        }


        private void CancelButton_Click(object sender, EventArgs e)
        {
            Cancel(this, new EventArgs());
        }

        private void ExportButton_Click(object sender, EventArgs e)
        {
            Export(this, new EventArgs());
        }

        private void ExcellAddInView_Load(object sender, EventArgs e)
        {

        }

        public int ViewingModeValue
        {
            get
            {
                    return this.ViewingModeComboBox.SelectedIndex;
            }
        }
        public int SelectedIndexValue
        {
            get
            {
                if (listBox1.SelectedItems.Count > 0)
                {
                    return this.listBox1.SelectedIndex;
                }
                else
                {
                    return 0;
                }
            }
        }

        public void SetupList(List<IDisplay> list)
        {
            listBox1.DataSource = list;
        }
        public void SetupGrid(List<IDisplay> list)
        {
            dataGridView1.DataSource = list;
        }



        private void ViewingMode_TextChanged(object sender, EventArgs e)
        {

        }

        private void ViewingModeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<IDisplay> list = new List<IDisplay>();
            IDisplay i = new Display();
            i.Title = " ";
            list.Add(i);
            listBox1.DataSource = list;
            listBox1.SelectedItem = " ";
            if (ViewingMode != null)
            {
                ViewingMode(this, new EventArgs());
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }

}
