//------------------------------------------------------------------------------
// <copyright file="SwapAllocation.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;

namespace AllocationSwapper
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class SwapAllocation
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("a9509730-4837-4d87-b727-3e5750ebf79d");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="SwapAllocation"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private SwapAllocation(Package package)
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
        public static SwapAllocation Instance
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
            Instance = new SwapAllocation(package);
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
            EnvDTE.DTE app = (EnvDTE.DTE)this.ServiceProvider.GetService(typeof(SDTE));
            if (app.ActiveDocument != null && app.ActiveDocument.Type == "Text")
            {
                EnvDTE.TextDocument text = (EnvDTE.TextDocument)app.ActiveDocument.Object(String.Empty);
                if (!text.Selection.IsEmpty)
                {
                    string newText = string.Empty;
                    var selectedText = text.Selection.Text.Replace(Environment.NewLine, "");
                    string[] textByLine = selectedText.Split(';');
                    foreach(string line in textByLine)
                    {
                        if (line.Contains("="))
                        {
                            string[] values = line.Split('=');
                            if(values.Length == 2)
                            {
                                newText += string.Format("{0} = {1};{2}", values[1], values[0], Environment.NewLine);
                            }
                        }
                        if (line == textByLine[textByLine.Length - 1])
                            newText = newText.TrimEnd(Environment.NewLine.ToCharArray());
                    }
                    text.Selection.Text = newText;
                    text.Selection.SmartFormat();
                }
            }
        }
    }
}
