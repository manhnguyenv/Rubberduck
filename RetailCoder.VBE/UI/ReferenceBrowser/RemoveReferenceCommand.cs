using System;
using NLog;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.ReferenceBrowser
{
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
}