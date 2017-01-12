using System;
using NLog;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.ReferenceBrowser
{
    public class AddReferenceFromFileCommand : CommandBase
    {
        private readonly IVBE _vbe;

        public AddReferenceFromFileCommand(IVBE vbe)
            : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            return !_vbe.ActiveVBProject.IsWrappingNullReference;
        }

        protected override void ExecuteImpl(object parameter)
        {
            var args = parameter as ProjectReferenceCommandParameters;
            if (args == null || string.IsNullOrEmpty(args.Path))
            {
                throw new ArgumentException();
            }

            var project = _vbe.ActiveVBProject;
            if (project == null)
            {
                throw new InvalidOperationException();
            }

            // throws COMException on failure
            args.Reference = project.References.AddFromFile(args.Path);
        }
    }
}