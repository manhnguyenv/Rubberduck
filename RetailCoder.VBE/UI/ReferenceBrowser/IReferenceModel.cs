using System;

namespace Rubberduck.UI.ReferenceBrowser
{
    public interface IReferenceModel
    {
        string FilePath { get; }
        string Name { get; }
        short MajorVersion { get; }
        short MinorVersion { get; }
        Guid Guid { get; }
    }
}