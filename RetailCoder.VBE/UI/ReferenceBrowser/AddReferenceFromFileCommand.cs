using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using NLog;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.ReferenceBrowser
{
    public class ProjectReferenceCommandParameters
    {
        private readonly string _path;

        public ProjectReferenceCommandParameters(string path)
        {
            _path = path;
        }

        public string Path { get { return _path; } }
        public Reference Reference { get; set; }
    }

    public class RemoveReferenceCommand : CommandBase
    {
        public RemoveReferenceCommand()
            : base(LogManager.GetCurrentClassLogger())
        {
            
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            return base.CanExecuteImpl(parameter);
        }

        protected override void ExecuteImpl(object parameter)
        {
            throw new NotImplementedException();
        }
    }

    public class AddReferenceFromFileCommand : CommandBase
    {
        private readonly VBE _vbe;

        public AddReferenceFromFileCommand(VBE vbe)
            : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
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