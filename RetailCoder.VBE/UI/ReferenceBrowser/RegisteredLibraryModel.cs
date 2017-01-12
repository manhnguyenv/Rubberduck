using Kavod.ComReflection;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.ReferenceBrowser
{
    public class RegisteredLibraryModel
    {
        private readonly string _guid;
        private readonly int _minorVersion;
        private readonly int _majorVersion;
        private readonly string _name;
        private readonly string _path;

        public RegisteredLibraryModel(IReference reference)
        {
            _path = reference.FullPath;
            _name = reference.Name;
            _majorVersion = reference.Major;
            _minorVersion = reference.Minor;
            _guid = reference.Guid;
        }

        internal RegisteredLibraryModel(LibraryRegistration library)
        {
            _path = library.FilePath;
            _name = library.Name;
            _majorVersion = library.MajorVersion;
            _minorVersion = library.MinorVersion;
            _guid = library.Guid.ToString();
        }

        public string FilePath { get { return _path; } }
        public string Name { get { return _name; } }
        public int MajorVersion { get { return _majorVersion; } }
        public int MinorVersion { get { return _minorVersion; } }
        public string Guid { get { return _guid; } }
    }
}