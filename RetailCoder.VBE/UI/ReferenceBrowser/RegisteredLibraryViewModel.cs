using System;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.ReferenceBrowser
{
    public class RegisteredLibraryViewModel : ViewModelBase, IComparable<RegisteredLibraryViewModel>
    {
        private readonly IVBProject _project;
        private readonly RegisteredLibraryModel _model;
        private readonly bool _isActiveReference;
        private readonly bool _canRemoveReference;

        public RegisteredLibraryViewModel(RegisteredLibraryModel model, IVBProject project)
        {
            _project = project;
            _model = model;

            IReference reference;
            _isActiveReference = TryGetProjectReference(_model.FilePath, out reference);
            _canRemoveReference = reference != null && !reference.IsBuiltIn;

            IsSelected = _isActiveReference;
        }

        public string FullPath { get { return _model.FilePath; } }
        public string Name { get { return _model.Name; } }
        public bool IsActiveProjectReference { get { return _isActiveReference; } }
        public bool CanRemoveReference { get { return _canRemoveReference; } }
        public string Guid { get { return _model.Guid; } }
        public int MinorVersion { get { return _model.MinorVersion; } }
        public int MajorVersion { get { return _model.MajorVersion; } }

        private bool _isSelected;
        public bool IsSelected
        {
            get { return _isSelected; }
            set
            {
                if (value == _isSelected)
                {
                    return;
                }
                _isSelected = value;
                OnPropertyChanged();
            }
        }

        public bool IsAdded { get { return _isSelected && !_isActiveReference; } }
        public bool IsRemoved { get { return !_isSelected && _isActiveReference; } }

        private bool TryGetProjectReference(string path, out IReference reference)
        {
            reference = _project.References.SingleOrDefault(item => item.FullPath == path);
            return reference != null;
        }

        #region IComparable

        public int CompareTo(RegisteredLibraryViewModel other)
        {
            if (CanRemoveReference && !other.CanRemoveReference)
            {
                return 1;
            }
            if (!CanRemoveReference && other.CanRemoveReference)
            {
                return -1;
            }
            if (IsActiveProjectReference && !other.IsActiveProjectReference)
            {
                return -1;
            }
            if (!IsActiveProjectReference && other.IsActiveProjectReference)
            {
                return 1;
            }
            return string.Compare(this.Name, other.Name, StringComparison.InvariantCultureIgnoreCase);
        }

        #endregion
    }
}