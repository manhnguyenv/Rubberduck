using System;
using Microsoft.Vbe.Interop;

namespace Rubberduck.UI.ReferenceBrowser
{
    public class VbaProjectReferenceModel : IReferenceModel
    {
        private readonly Reference _reference;

        public VbaProjectReferenceModel(Reference reference)
        {
            _reference = reference;
        }

        public string FilePath { get { return _reference.FullPath; } }
        public string Name { get { return _reference.Name; } }
        public short MajorVersion { get { return (short) _reference.Major; } }
        public short MinorVersion { get { return (short) _reference.Minor; } }
        public Guid Guid { get { return Guid.Parse(_reference.Guid); } }
    }
}