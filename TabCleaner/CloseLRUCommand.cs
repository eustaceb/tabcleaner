using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace TabCleaner
{

    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class CloseLRUCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0101;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("914e0122-bbe2-4687-bf51-8f4997603c42");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="CloseLRUCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private CloseLRUCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static CloseLRUCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new CloseLRUCommand(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            // Get the list of open documents
            var dteService = ServiceProvider.GetService(typeof(DTE)) as DTE;
            var openDocs = dteService.Documents;

            // Get the package service
            var tabCleanerPkg = ServiceProvider.GetService(typeof(TabCleanerPackage)) as TabCleanerPackage;

            // Pull out the least recently used documents
            var LRUDocs = tabCleanerPkg.GetLRUDocs();

            // Match against every open document against the ones
            // @TODO: What happens with multiple windows of the same document?
            foreach (Document doc in openDocs)
            {
                foreach (string docPath in LRUDocs)
                {
                    if (doc.FullName.Equals(docPath))
                    {
                        if (tabCleanerPkg.CloseModified || doc.Saved)
                            doc.Close();
                    }
                }
            }
        }
    }
}
