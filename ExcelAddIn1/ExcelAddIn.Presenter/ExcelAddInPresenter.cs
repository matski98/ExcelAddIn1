using ExcelAddIn.Model.Interfaces;
using ExcelAddIn.GUI.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelAddIn.Model;
using ExcelAddIn.Application;
using ExcelAddIn.Framework;

namespace ExcelAddIn.Presenter
{
    public class ExcelAddInPresenter
    {
        private IExcelAddInModel _model;
        private IExcelAddInView _view;
        private List<IDisplay> _displaylistview=new List<IDisplay>();
        private List<IDisplay> _displaysgridview = new List<IDisplay>();
        public ExcelAddInPresenter(IExcelAddInView view, IExcelAddInModel model)
        {
            _view = view;
            _model = model;
            _view.Cancel += _view_Cancel;
            _view.Export += _view_Export;
            _view.ViewingMode += _view_ModeChanged;
            _view.SelectedIndex += _view_SelectedIndex;
            _displaylistview = PersonsTransformListBox(_model);
            if (_model.Persons.Count!=0)
                _displaysgridview = PersonsTransformGridView(_model, _displaylistview[_view.SelectedIndexValue].ToString());
            this._view.SetupList(_displaylistview);
            this._view.SetupGrid(_displaysgridview);


            _view.Showdialog();

        }
        private void _view_SelectedIndex(object sender, EventArgs e)
        {
            if (_view.ViewingModeValue == 0)
            {
                if (_model.Persons.Count != 0)
                    _displaysgridview = PersonsTransformGridView(_model, _displaylistview[_view.SelectedIndexValue].ToString());
                this._view.SetupGrid(_displaysgridview);
            }
            else if (_view.ViewingModeValue == 1)
            {
                if (_model.Persons.Count != 0)
                    _displaysgridview = ProjectsTransformGridView(_model, _displaylistview[_view.SelectedIndexValue].ToString());
                this._view.SetupGrid(_displaysgridview);
            }
            else
            {
                _displaysgridview = PersonsTransformGridView(_model, _displaylistview[_view.SelectedIndexValue].ToString());
                this._view.SetupGrid(_displaysgridview);
            }
        }

        private void _view_Cancel(object sender, EventArgs e)
        {
            _view.Close();
        }
        private void _view_Export(object sender, EventArgs e)
        {
            IExportExcel service = ServiceProvider.GetService<IExportExcel>();
            service.Export(_model);
        }
        private List<IDisplay> PersonsTransformGridView(IExcelAddInModel model, string czlowiek)
        {
            List <IDisplay> persons = new List<IDisplay>();
            foreach(IPerson person in model.Persons)
            {
                if(string.Format("{0} {1}", person.Name, person.Surname) == czlowiek)
                {
                    foreach (IProject project in person.Projects)
                    {
                        IDisplay display = new Display();
                        display.Title = project.Name;
                        display.Hours = project.Number;
                        persons.Add(display);
                    }
                }
                
            }
            return persons;
        }
        private List<IDisplay> ProjectsTransformGridView(IExcelAddInModel model, string projekt)
        {
            List<IDisplay> displays = new List<IDisplay>();
            foreach (IPerson person in model.Persons)
            {
                foreach (IProject project in person.Projects)
                {
                    if (project.Name == projekt)
                    {
                        IDisplay display = new Display();
                        display.Title = string.Format("{0} {1}", person.Name, person.Surname);
                        display.Hours = project.Number;
                        displays.Add(display);
                    }
                }
            }
            return displays;
        }
        private List<IDisplay> PersonsTransformListBox(IExcelAddInModel model)
        {
            List<string> persons = new List<string>();
            foreach(IPerson person in model.Persons)
            {
                
                persons.Add(string.Format("{0} {1}", person.Name, person.Surname));
            }
            persons.Sort();
            List<IDisplay> displays = new List<IDisplay>();
            foreach(string person_name in persons)
            {
                IDisplay display = new Display();
                display.Title = person_name;
                displays.Add(display);
            }
            return displays;
        }
        private List<IDisplay> ProjectsTransformListBox(IExcelAddInModel model)
        {
            List<string> projects = new List<string>();
            foreach (IPerson person in model.Persons)
            {
                foreach (IProject project in person.Projects)
                {
                    if (!projects.Contains(project.Name))
                    {
                        projects.Add(project.Name);
                    }
                }

            }
            projects.Sort();
            List<IDisplay> displays = new List<IDisplay>();
            foreach (string project_name in projects)
            {
                IDisplay display = new Display();
                display.Title = project_name;
                displays.Add(display);
            }
            return displays;
        }

        private void _view_ModeChanged(object sender, EventArgs e)
        {
            if (_view.ViewingModeValue == 0)
            {
                _displaylistview = PersonsTransformListBox(_model);
                if (_model.Persons.Count != 0)
                    _displaysgridview = PersonsTransformGridView(_model, _displaylistview[_view.SelectedIndexValue].ToString());
                this._view.SetupList(_displaylistview);
                this._view.SetupGrid(_displaysgridview);
            }
            else if (_view.ViewingModeValue == 1)
            {
                _displaylistview = ProjectsTransformListBox(_model);
                if (_model.Persons.Count != 0)
                    _displaysgridview = ProjectsTransformGridView(_model, _displaylistview[_view.SelectedIndexValue].ToString());
                this._view.SetupList(_displaylistview);
                this._view.SetupGrid(_displaysgridview);
            }
        }
    }
}
