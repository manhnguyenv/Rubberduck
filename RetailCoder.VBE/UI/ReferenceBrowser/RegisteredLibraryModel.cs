using System;
using Kavod.ComReflection;

namespace Rubberduck.UI.ReferenceBrowser
{
    public class RegisteredLibraryModel : IReferenceModel
    {
        private readonly LibraryRegistration _library;

        internal RegisteredLibraryModel(LibraryRegistration library)
        {
            _library = library;
        }

        public string FilePath { get { return _library.FilePath; } }
        public string Name { get { return _library.Name; } }
        public short MajorVersion { get { return _library.MajorVersion; } }
        public short MinorVersion { get { return _library.MinorVersion; } }
        public Guid Guid { get { return _library.Guid; } }
    }
}