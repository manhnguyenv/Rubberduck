using System;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.AddRemoveReferences
{
    public class ReferenceModel : ViewModelBase
    {
        public ReferenceModel(IVBProject project)
        {
            Name = project.Name;
            Guid = string.Empty;
            Description = project.Description;
            Version = new Version();
            FullPath = project.FileName;
            IsBuiltIn = false;
            Type = ReferenceKind.Project;
            IsVisible = true;
        }

        public ReferenceModel(RegisteredLibraryInfo info)
        {
            Name = info.Name;
            Guid = info.Guid;
            Description = info.Description;
            Version = info.Version;
            FullPath = info.FullPath;
            IsBuiltIn = false;
            Type = ReferenceKind.TypeLibrary;
            Flags = info.Flags;
            SubKey = info.SubKey;
            IsVisible = true;
        }

        public ReferenceModel(ComProject reference)
        {
            Name = reference.Documentation.Name;
            Guid = reference.Guid.ToString();
            Description = reference.Documentation.DocString;
            Version = new Version((int)reference.MajorVersion, (int)reference.MinorVersion);
            FullPath = reference.Path;
            IsBuiltIn = false;
            Type = ReferenceKind.TypeLibrary;
            IsVisible = true;
            IsSelected = true;
        }

        public ReferenceModel(IReference reference, int priority)
        {
            IsSelected = true;
            Priority = priority;
            Name = reference.Name;
            Guid = reference.Guid;
            Description = reference.Description;
            Version = new Version(reference.Major, reference.Minor);
            FullPath = reference.FullPath;
            IsBuiltIn = reference.IsBuiltIn;
            IsBroken = reference.IsBroken;
            Type = reference.Type;
            IsVisible = true;
        }

        private bool _isVisible;
        public bool IsVisible
        {
            get => _isVisible;
            set
            {
                if (_isVisible != value)
                {
                    _isVisible = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged();
                }
            }
        }

        private int _priority;
        public int Priority
        {
            get => _priority;
            set
            {
                if (_priority != value)
                {
                    _priority = value;
                    OnPropertyChanged();
                }
            }
        }

        public string Name { get; }
        public string Guid { get; }
        public string Description { get; }
        public Version Version { get; }
        public string FullPath { get; }
        public bool IsBuiltIn { get; }
        public bool IsBroken { get; }
        public int Flags { get; }
        public int SubKey { get; }
        public ReferenceKind Type { get; }

        public ReferenceStatus Status => IsBuiltIn
            ? ReferenceStatus.BuiltIn
            : IsBroken
                ? ReferenceStatus.Broken
                : IsSelected
                    ? ReferenceStatus.Loaded
                    : ReferenceStatus.None;

        public override bool Equals(object obj)
        {
            var other = obj as ReferenceModel;
            if (other == null) { return false; }

            return other.Guid == Guid 
                && other.Version == Version 
                && other.FullPath == FullPath;
        }

        public override int GetHashCode()
        {
            return HashCode.Compute(Guid, Version, FullPath);
        }
    }
}