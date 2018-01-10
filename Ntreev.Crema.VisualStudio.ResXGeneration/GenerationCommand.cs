using System;
using System.Linq;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;
using System.Diagnostics;
using System.ComponentModel.Composition.Hosting;
using Ntreev.Crema.RuntimeService;
using Ntreev.Crema.Runtime.Serialization;
using Ntreev.Crema.Data;
using Ntreev.Library.IO;
using System.Resources;
using System.Collections;
using System.Collections.Generic;
using Ntreev.Library;
using System.IO;
using Microsoft.CSharp;
using System.Resources.Tools;
using System.CodeDom.Compiler;
using Ntreev.Crema.VisualStudio.ResXGeneration.Dialogs.Views;
using Ntreev.Crema.VisualStudio.ResXGeneration.Dialogs.ViewModels;

namespace Ntreev.Crema.VisualStudio.ResXGeneration
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class GenerationCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int ItemCommandId = 0x0100;
        public const int ProjectCommandId = 0x0101;
        public const int SolutionCommandId = 0x0102;

        public const int DebugCommandId = 0x0112;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("0b83c475-7412-46a8-ac9f-981413375777");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        private CompositionContainer container;

        /// <summary>
        /// Initializes a new instance of the <see cref="GenerationCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private GenerationCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            if (this.ServiceProvider.GetService(typeof(IMenuCommandService)) is OleMenuCommandService commandService)
            {
                var itemMenuCommandID = new CommandID(CommandSet, ItemCommandId);
                var itemMenuItem = new OleMenuCommand(this.ItemMenuItemCallback, itemMenuCommandID);
                commandService.AddCommand(itemMenuItem);
                itemMenuItem.BeforeQueryStatus += ItemMenuItem_BeforeQueryStatus;

                var projectMenuCommandID = new CommandID(CommandSet, ProjectCommandId);
                var projectMenuItem = new OleMenuCommand(this.ProjectMenuItemCallback, projectMenuCommandID);
                commandService.AddCommand(projectMenuItem);
                itemMenuItem.BeforeQueryStatus += ProjectMenuItem_BeforeQueryStatus;

                var solutionMenuCommandID = new CommandID(CommandSet, SolutionCommandId);
                var solutionMenuItem = new OleMenuCommand(this.SolutionMenuItemCallback, solutionMenuCommandID);
                commandService.AddCommand(solutionMenuItem);

                var debugMenuCommandID = new CommandID(CommandSet, DebugCommandId);
                var debugMenuItem = new OleMenuCommand(this.DebugMenuItemCallback, debugMenuCommandID);
                commandService.AddCommand(debugMenuItem);
            }

            this.container = new CompositionContainer(new AssemblyCatalog(typeof(IRuntimeService).Assembly));
        }

        private void ItemMenuItem_BeforeQueryStatus(object sender, EventArgs e)
        {
            if (sender is OleMenuCommand command)
            {
                command.Enabled = this.IsEmbeddedResources();
            }
        }

        private void ProjectMenuItem_BeforeQueryStatus(object sender, EventArgs e)
        {
            if (sender is OleMenuCommand command)
            {

            }
        }

        private bool IsEmbeddedResources()
        {
            var dte = this.ServiceProvider.GetService(typeof(SDTE)) as EnvDTE80.DTE2;
            if (dte.SelectedItems.Count == 0)
                return false;

            dte.Solution.GetAddress();

            for (var i = 1; i <= dte.SelectedItems.Count; i++)
            {
                var item = dte.SelectedItems.Item(i);
                if (item == null || item.ProjectItem == null)
                    return false;

                var projectItem = item.ProjectItem;
                if (projectItem.IsEmbeddedResource() == false)
                    return false;
            }

            return true;
        }

        private void Print(ProjectItem projectItem)
        {
            for (var i = 1; i <= projectItem.Properties.Count; i++)
            {
                var item = projectItem.Properties.Item(i);
                //if (item.Name == "ItemType") // EmbeddedResource
                Trace.WriteLine($"{item.Name}: {item.Value}");

                //IsDependentFile
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static GenerationCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        internal IServiceProvider ServiceProvider
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
            Instance = new GenerationCommand(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void ItemMenuItemCallback(object sender, EventArgs e)
        {
            var runtimeService = this.container.GetExportedValue<IRuntimeService>();
            var settings = this.Solution.GetSettings();
            var data = runtimeService.GetDataGenerationData(settings.Address, settings.DataBase, $"{TagInfo.All}", null, false, -1);
            var dataSet = SerializationUtility.Create(data);
            var projectInfoTable = dataSet.Tables[settings.ProjectInfo];

            var successCount = 0;
            var failCount = 0;
            GenerationOutput.WriteLine("------ Start To Update Crema Resources: Selected Resource Files ------");

            foreach (var item in this.GetSelectedItems())
            {
                try
                {
                    item.Write(projectInfoTable);
                    GenerationOutput.WriteLine($" >{item.GetFullPath()}");
                    successCount++;
                }
                catch (Exception ex)
                {
                    GenerationOutput.WriteLine($"X>{item.GetFullPath()}:  error: {ex.Message}");
                    failCount++;
                }
            }

            GenerationOutput.WriteLine($"========== Update: Success {successCount}, Fail {failCount} : {DateTime.Now} ==========");
            GenerationOutput.WriteLine();
        }

        private void ProjectMenuItemCallback(object sender, EventArgs e)
        {
            var runtimeService = this.container.GetExportedValue<IRuntimeService>();
            var settings = this.Solution.GetSettings();
            var data = runtimeService.GetDataGenerationData(settings.Address, settings.DataBase, $"{TagInfo.All}", null, false, -1);
            var dataSet = SerializationUtility.Create(data);
            var projectInfoTable = dataSet.Tables[settings.ProjectInfo];

            foreach (var item in this.GetSelectedProjects())
            {
                item.Write(projectInfoTable);
            }
            GenerationOutput.WriteLine();
        }

        private void SolutionMenuItemCallback(object sender, EventArgs e)
        {
            var runtimeService = this.container.GetExportedValue<IRuntimeService>();
            var settings = this.Solution.GetSettings();
            var data = runtimeService.GetDataGenerationData(settings.Address, settings.DataBase, $"{TagInfo.All}", null, false, -1);
            var dataSet = SerializationUtility.Create(data);
            var projectInfoTable = dataSet.Tables[settings.ProjectInfo];


            foreach (var item in this.Solution)
            {
                int qwer = 0;
            }
            //for (var i = 1; i <= this.Solution.Projects.Count; i++)
            //{
            //    var item = this.Solution.Projects.Item(i);
            //    item.Write(projectInfoTable);
            //}
            //GenerationOutput.WriteLine();
        }

        private void DebugMenuItemCallback(object sender, EventArgs e)
        {
            foreach (var item in this.GetSelectedItems())
            {
                item.PrintProperties();
            }
        }

        //private void Write(ProjectItem projectItem, CremaDataTable projectInfoTable)
        //{
        //    if (this.FindTable(projectInfoTable, projectItem) is CremaDataTable itemTable)
        //    {
        //        try
        //        {
        //            projectItem.Write(itemTable);
        //        }
        //        catch (Exception ex)
        //        {
        //            VsShellUtilities.ShowMessageBox(this.ServiceProvider, ex.Message, "Error",
        //                OLEMSGICON.OLEMSGICON_CRITICAL, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        //        }
        //    }
        //}

        //private CremaDataTable FindTable(CremaDataTable projectInfoTable, ProjectItem projectItem)
        //{
        //    var project = projectItem.ContainingProject;
        //    var projectPath = project.GetLocalPath();
        //    var projectItemPath = projectItem.GetLocalPathWithoutCultureInfo();
        //    var dataSet = projectInfoTable.DataSet;

        //    foreach (var item in projectInfoTable.Rows)
        //    {
        //        if (item["ProjectPath"] is string text && text == projectPath)
        //        {
        //            if (GetExportInfo(item) is CremaDataRow exportInfoRow)
        //            {
        //                return dataSet.Tables[exportInfoRow["TableName"] as string];
        //            }
        //        }
        //    }

        //    return null;

        //    CremaDataRow GetExportInfo(CremaDataRow dataRow)
        //    {
        //        foreach (var item in dataRow.GetChildRows("ExportInfo"))
        //        {
        //            if (item["FileName"] is string text && text == projectItemPath)
        //            {
        //                return item;
        //            }
        //        }
        //        return null;
        //    }
        //}

        private IEnumerable<ProjectItem> GetSelectedItems()
        {
            var dte = this.ServiceProvider.GetService(typeof(SDTE)) as EnvDTE80.DTE2;
            if (dte.SelectedItems.Count > 0)
            {
                for (var i = 1; i <= dte.SelectedItems.Count; i++)
                {
                    var item = dte.SelectedItems.Item(i);
                    if (item != null || item.ProjectItem != null)
                        yield return item.ProjectItem;
                }
            }
        }

        private IEnumerable<Project> GetSelectedProjects()
        {
            var dte = this.ServiceProvider.GetService(typeof(SDTE)) as EnvDTE80.DTE2;
            if (dte.SelectedItems.Count > 0)
            {
                for (var i = 1; i <= dte.SelectedItems.Count; i++)
                {
                    var item = dte.SelectedItems.Item(i);
                    if (item != null || item.Project != null)
                        yield return item.Project;
                }
            }
        }

        private Solution Solution
        {
            get
            {
                if (this.ServiceProvider.GetService(typeof(SDTE)) is EnvDTE80.DTE2 dte)
                {
                    return dte.Solution;
                }
                throw new InvalidOperationException();
            }
        }
    }
}
