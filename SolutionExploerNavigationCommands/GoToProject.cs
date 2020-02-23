using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace SolutionExploerNavigationCommands
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class GoToProject
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("7b2d2438-29f6-4c6e-aade-25226c880029");
        private static DTE2 _dte;

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage _package;

        /// <summary>
        /// Initializes a new instance of the <see cref="GoToProject"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private GoToProject(AsyncPackage package, OleMenuCommandService commandService, DTE2 dte)
        {
            this._package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));
            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new SolutionExplorerNavigationCommand(Execute, menuCommandID, dte);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static GoToProject Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this._package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in GoToProject's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            _dte = await package.GetServiceAsync(typeof(DTE)) as DTE2;
            Assumes.Present(_dte);
            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new GoToProject(package, commandService, _dte as DTE2);
        }

        /// <summary> 
        /// This function is the callb ack used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_dte.ActiveWindow.Type != vsWindowType.vsWindowTypeSolutionExplorer)
            {
                return;
            }
            var items = _dte.ToolWindows.SolutionExplorer.SelectedItems as UIHierarchyItem[];
            if (items?.Length != 1)
            {
                return;
            }
            var item = items[0].Collection.Parent as UIHierarchyItem;
            while (item != null)
            {
                if (item.Object is Project || item.Object is Solution)
                {
                    break;
                }
                item = item.Collection.Parent as UIHierarchyItem;
            }
            item?.Select(vsUISelectionType.vsUISelectionTypeSelect);
        }
    }

    public class SolutionExplorerNavigationCommand : MenuCommand
    {
        private readonly DTE2 _dte;

        public SolutionExplorerNavigationCommand(EventHandler handler, CommandID commandID, DTE2 dte) : base(handler, commandID)
        {
            _dte = dte;
        }

        public override bool Enabled
        {
            get
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                return _dte.ToolWindows.SolutionExplorer.SelectedItems is UIHierarchy[] array && array.Length == 1;
            }

            set
            {
                base.Enabled = value;
            }
        }
    }

}
