using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Input;
using NLog;
using Rubberduck.AddRemoveReferences;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.AddRemoveReferences
{
    public class AddRemoveReferencesViewModel : ViewModelBase
    {
        private static readonly ILogger Logger = LogManager.GetCurrentClassLogger();
        private readonly IProjectReferencesService _service;
        private readonly IOpenFileDialog _fileDialog;
        private readonly IMessageBox _messageBox;

        public AddRemoveReferencesViewModel(IRegisteredLibraryFinderService finder, IProjectReferencesService service, IOpenFileDialog fileDialog, IMessageBox messageBox)
        {
            _comLibraries = new ObservableCollection<ReferenceModel>(finder.FindRegisteredLibraries());
            _vbProjects = service.AvailableProjects;
            _service = service;
            _fileDialog = fileDialog;
            _messageBox = messageBox;
            _fileDialog.Title = RubberduckUI.ResourceManager.GetString("AddRemoveReferences_BrowseTitle");
            _fileDialog.Filter = RubberduckUI.ResourceManager.GetString("AddRemoveReferences_BrowseFilter");
            _fileDialog.CheckFileExists = true;
            _fileDialog.Multiselect = true;

            CancelCommand = new DelegateCommand(Logger, o => OnCancel(), o => true);
            OkCommand = new DelegateCommand(Logger, o => OnAccept(), o => true);
            ToggleSelectedCommand = new DelegateCommand(Logger, ExecuteToggleSelectedCommand, CanExecuteToggleSelectedCommand);
            MoveUpCommand = new DelegateCommand(Logger, ExecuteMoveUpCommand, CanExecuteMoveUpCommand);
            MoveDownCommand = new DelegateCommand(Logger, ExecuteMoveDownCommand, CanExecuteMoveDownCommand);
            BrowseCommand = new DelegateCommand(Logger, ExecuteBrowseCommand, o => true);

            RefreshModel();
        }

        private void RefreshModel()
        {
            var model = new HashSet<ReferenceModel>();
            foreach (var projectRef in _service.References)
            {
                model.Add(projectRef);
            }
            foreach (var library in _comLibraries)
            {
                library.IsSelected = model.Contains(library);
            }
        }

        public event EventHandler Cancelled;
        public event EventHandler Accepted;

        private void OnCancel()
        {
            Cancelled?.Invoke(this, EventArgs.Empty);
        }

        private void OnAccept()
        {
            Accepted?.Invoke(this, EventArgs.Empty);
        }

        /// <summary>
        /// Prompts user for a .tlb, .dll, or .ocx file, and attempts to append it to <see cref="ProjectReferences"/>.
        /// </summary>
        public ICommand BrowseCommand { get; }

        /// <summary>
        /// Cancels all changes and closes the dialog.
        /// </summary>
        public ICommand CancelCommand { get; }

        /// <summary>
        /// Applies all changes and closes the dialog.
        /// </summary>
        public ICommand OkCommand { get; }

        /// <summary>
        /// Applies all changes, without closing the dialog.
        /// </summary>
        public ICommand ApplyCommand { get; }

        private bool _isDirty;

        public bool IsDirty
        {
            get => _isDirty;
            set
            {
                if (_isDirty != value)
                {
                    _isDirty = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Toggles whether the selected reference is included or not.
        /// </summary>
        public ICommand ToggleSelectedCommand { get; }

        /// <summary>
        /// Moves the <see cref="SelectedReference"/> up on the 'Priority' tab.
        /// </summary>
        public ICommand MoveUpCommand { get; }

        /// <summary>
        /// Moves the <see cref="SelectedReference"/> down on the 'Priority' tab.
        /// </summary>
        public ICommand MoveDownCommand { get; }

        private readonly ObservableCollection<ReferenceModel> _comLibraries;
        public IEnumerable<ReferenceModel> ComLibraries { get { return _comLibraries.Where(r => r.IsSelected || r.IsVisible); } }

        private readonly IEnumerable<ReferenceModel> _vbProjects;
        public IEnumerable<ReferenceModel> VbaProjects { get { return _vbProjects.Where(r => r.IsSelected || r.IsVisible); } }

        private ReferenceModel _selectedLibrary;
        public ReferenceModel SelectedLibrary
        {
            get => _selectedLibrary;
            set
            {
                if (_selectedLibrary == null || !_selectedLibrary.Equals(value))
                {
                    _selectedLibrary = value;
                    OnPropertyChanged();
                }
            }
        }

        private ReferenceModel _selectedProject;
        public ReferenceModel SelectedProject
        {
            get => _selectedProject;
            set
            {
                if (_selectedProject == null || !_selectedProject.Equals(value))
                {
                    _selectedProject = value;
                    OnPropertyChanged();
                }
            }
        }

        private ReferenceModel _selectedReference;
        public ReferenceModel SelectedReference
        {
            get => _selectedReference;
            set
            {
                if (_selectedReference == null || !_selectedReference.Equals(value))
                {
                    _selectedReference = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Gets all selected COM libraries and VBA projects, sorted by priority.
        /// </summary>
        public IOrderedEnumerable<ReferenceModel> ProjectReferences
        {
            get
            {
                return ComLibraries.Where(r => r.IsSelected)
                    .Union(VbaProjects.Where(r => r.IsSelected))
                    .OrderBy(r => r.Priority);
            }
        }

        private bool CanExecuteToggleSelectedCommand(object item)
        {
            var model = (ReferenceModel)item;
            return !model.IsBuiltIn;
        }

        private void ExecuteToggleSelectedCommand(object item)
        {
            var model = (ReferenceModel)item;
            model.IsSelected = !model.IsSelected;
        }

        private bool CanExecuteMoveUpCommand(object item)
        {
            var model = (ReferenceModel)item;
            return !model.IsBuiltIn && model.Priority > ProjectReferences.First().Priority;
        }

        private void ExecuteMoveUpCommand(object item)
        {
            var model = (ReferenceModel)item;
            var allRefs = ProjectReferences.ToList();
            var oldIndex = allRefs.IndexOf(model);
            allRefs.RemoveAt(oldIndex);
            allRefs.Insert(oldIndex - 1, model);
            for (var priority = 0; priority < allRefs.Count; priority++)
            {
                allRefs[priority].Priority = priority;
            }

            IsDirty = true;
        }

        private bool CanExecuteMoveDownCommand(object item)
        {
            var model = (ReferenceModel)item;
            return !model.IsBuiltIn && model.Priority < ProjectReferences.Last().Priority;
        }

        private void ExecuteMoveDownCommand(object item)
        {
            var model = (ReferenceModel)item;
            var allRefs = ProjectReferences.ToList();
            var oldIndex = allRefs.IndexOf(model);
            allRefs.RemoveAt(oldIndex);
            allRefs.Insert(oldIndex + 1, model);
            for (var priority = 0; priority < allRefs.Count; priority++)
            {
                allRefs[priority].Priority = priority;
            }

            IsDirty = true;
        }

        private void ExecuteBrowseCommand(object o)
        {
            if (_fileDialog.ShowDialog() != DialogResult.Cancel)
            {
                foreach (var fileName in _fileDialog.FileNames)
                {
                    try
                    {
                        var info = _service.GetLibraryInfo(/*System.IO.Path.Combine(*//*path?*/fileName/*)*/);
                        if (info != null)
                        {
                            _comLibraries.Add(info);
                            IsDirty = true;
                        }
                        else
                        {
                            _messageBox.Show(RubberduckUI.AddRemoveReferences_CouldNotLoadLibrary,
                                RubberduckUI.ToolsMenu_AddRemoveReferences, 
                                MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                    }
                    catch (Exception exception)
                    {
                        Logger.Error(exception);
                        _messageBox.Show(exception.Message, RubberduckUI.ToolsMenu_AddRemoveReferences, 
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
    }
}
