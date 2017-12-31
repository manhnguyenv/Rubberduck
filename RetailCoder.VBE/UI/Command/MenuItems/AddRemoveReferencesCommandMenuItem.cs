using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class AddRemoveReferencesCommandMenuItem : CommandMenuItemBase
    {
        public AddRemoveReferencesCommandMenuItem(CommandBase command) : base(command)
        {             
        }

        public override string Key => "ToolsMenu_AddRemoveReferences";

        public override int DisplayOrder => (int) ToolsMenuItemDisplayOrder.ProjectReferences;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return Command.CanExecute(null);
        }
    }
}