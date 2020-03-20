using EnvDTE;
using Ntreev.Crema.Data;
using Ntreev.Library;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace Ntreev.Crema.VisualStudio.ResXGeneration
{
    static class ProjectExtensions
    {
        public static string GetFileName(this Project project)
        {
            return project.FileName;
        }

        public static string GetFullName(this Project project)
        {
            return project.FullName;
        }

        /// <summary>
        /// 솔루션 경로 기준으로 계산된 파일 경로
        /// </summary>
        public static string GetLocalPath(this Project project)
        {
            var solution = project.DTE.Solution;
            return UriUtility.MakeRelative(solution.FullName, project.FullName);
        }

        public static ProjectItem GetProjectItem(this Project project, string name)
        {
            for (var i = 1; i <= project.ProjectItems.Count; i++)
            {
                var projectItem = project.ProjectItems.Item(i);
                if (projectItem.Name == name)
                    return projectItem;
            }
            return null;
        }

        public static bool ContainsProjectItem(this Project project, string name)
        {
            for (var i = 1; i <= project.ProjectItems.Count; i++)
            {
                var projectItem = project.ProjectItems.Item(i);
                if (projectItem.Name == name)
                    return true;
            }
            return false;
        }

        public static string GetRootNamespace(this Project project)
        {
            return Path.GetFileNameWithoutExtension(project.GetFullName());
        }

        //public static void Write(this Project project, CremaDataTable projectInfoTable)
        //{
        //    var items = CollectProjectItems(project);

        //    var successCount = 0;
        //    var failCount = 0;

        //    GenerationOutput.WriteLine($"------ Update from {settings.Address}.{settings.DataBase}.{settings.ProjectInfo}: Project: {project.Name} ------");
        //    foreach (var item in project.CollectProjectItems())
        //    {
        //        try
        //        {
        //            item.Write(projectInfoTable);
        //            GenerationOutput.WriteLine($"O>{item.GetFullPath()}");
        //            successCount++;
        //        }
        //        catch (Exception ex)
        //        {
        //            GenerationOutput.WriteLine($" >{item.GetFullPath()}:  error: {ex.Message}");
        //            failCount++;
        //        }
        //    }
        //    GenerationOutput.WriteLine($"========== Update: Success {successCount}, Fail {failCount}: {DateTime.Now} ==========");
        //}

        public static IEnumerable<ProjectItem> GetProjectItems(this Project project)
        {
            if (project.ProjectItems != null)
            {
                for (var i = 1; i <= project.ProjectItems.Count; i++)
                {
                    var projectItem = project.ProjectItems.Item(i);
                    if (projectItem.IsEmbeddedResource() == true)
                        yield return projectItem;

                    foreach (var item in CollectProjectItems(projectItem))
                    {
                        yield return item;
                    }
                }
            }
        }

        private static IEnumerable<ProjectItem> CollectProjectItems(this ProjectItem projectItem)
        {
            if (projectItem.ProjectItems != null)
            {
                for (var i = 1; i <= projectItem.ProjectItems.Count; i++)
                {
                    var item = projectItem.ProjectItems.Item(i);
                    if (item.IsEmbeddedResource() == true)
                        yield return item;

                    foreach (var subItem in CollectProjectItems(item))
                    {
                        yield return subItem;
                    }
                }
            }
        }
    }
}
