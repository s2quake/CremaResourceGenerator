using EnvDTE;
using Ntreev.Library.Serialization;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Ntreev.Crema.VisualStudio.ResXGeneration
{
    static class SolutionExtensions
    {
        //public const string lkj;

        public static string GetFullName(this Solution solution)
        {
            return solution.FullName;
        }

        public static Project GetProject(this Solution solution, string name)
        {
            for (var i = 1; i <= solution.Count; i++)
            {
                var project = solution.Item(i);
                if (project.Name == name)
                {
                    return project;
                }
            }
            return null;
        }

        public static bool ContainsProject(this Solution solution, string name)
        {
            for (var i = 1; i <= solution.Count; i++)
            {
                var project = solution.Item(i);
                if (project.Name == name)
                {
                    return true;
                }
            }
            return false;
        }

        public static Project GetSelectedProject(this Solution solution)
        {
            for (var i = 1; i <= solution.Count; i++)
            {
                var project = solution.Item(i);
                //if (project.)
                //{
                //    return project;
                //}
                int qewr = 0;
            }
            return null;
        }

        public static string GetAddress(this Solution solution)
        {
            if (solution.ContainsSettings() == true)
            {
                return solution.GetSettings().Address;
            }
            return string.Empty;
        }

        public static string GetDataBase(this Solution solution)
        {
            if (solution.ContainsSettings() == true)
            {
                return solution.GetSettings().DataBase;
            }
            return string.Empty;
        }

        private static bool ContainsSettings(this Solution solution)
        {
            if (solution.ContainsProject("CremaResX") == false)
                return false;
            var project = solution.GetProject("CremaResX");
            if (project.ContainsProjectItem("settings.xml") == false)
                return false;
            return true;
        }

        public static ResXInfo GetSettings(this Solution solution)
        {
            var fullName = solution.GetFullName();
            var xmlPath = fullName + ".cremaresx";

            if (File.Exists(xmlPath) == true)
            {
                if (XmlSerializerUtility.Read(xmlPath, typeof(ResXInfo)) is ResXInfo obj)
                {
                    return obj;
                }
            }

            throw new InvalidOperationException();
        }

        public static IEnumerable<Project> CollectProjects(this Solution solution)
        {
            for (var i = 1; i <= solution.Projects.Count; i++)
            {
                var item = solution.Projects.Item(i);
                yield return item;
            }
        }
    }
}
