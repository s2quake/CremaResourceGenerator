using EnvDTE;
using Microsoft.CSharp;
using Ntreev.Crema.Data;
using Ntreev.Library;
using Ntreev.Library.IO;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Resources;
using System.Resources.Tools;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace Ntreev.Crema.VisualStudio.ResXGeneration
{
    static class ProjectItemExtensions
    {
        public const string EmbeddedResourceType = "EmbeddedResource";
        public const string CustomTool = "ResXFileCodeGenerator";

        public static bool GetIsDependentFile(this ProjectItem projectItem)
        {
            if (projectItem.GetProperty("IsDependentFile") is bool value)
                return value;
            return false;
        }

        public static string GetItemType(this ProjectItem projectItem)
        {
            if (projectItem.GetProperty("ItemType") is string value)
                return value;
            return string.Empty;
        }

        // ResXFileCodeGenerator
        public static string GetCustomTool(this ProjectItem projectItem)
        {
            if (projectItem.GetProperty("CustomTool") is string value)
                return value;
            return string.Empty;
        }

        /// <summary>
        /// 리소스 디자이너 파일명
        /// </summary>
        public static string GetCustomToolOutput(this ProjectItem projectItem)
        {
            if (projectItem.GetProperty("CustomToolOutput") is string value)
            {
                var directoryPath = Path.GetDirectoryName(projectItem.GetFullPath());
                return Path.Combine(directoryPath, value);
            }
            return string.Empty;
        }

        /// <summary>
        /// 프로젝트 경로 기준으로 계산된 파일 경로
        /// </summary>
        public static string GetLocalPath(this ProjectItem projectItem)
        {
            var project = projectItem.ContainingProject;
            return UriUtility.MakeRelative(project.FullName, projectItem.GetFullPath());
        }

        public static string GetLocalPathWithoutCultureInfo(this ProjectItem projectItem)
        {
            var localPath = GetLocalPath(projectItem);

            var match = Regex.Match(localPath, "(.+)([.][^.]+)([.]resx)$");
            if (match.Success == true)
            {
                var group = match.Groups[2];
                localPath = localPath.Remove(group.Index, group.Length);
            }

            return localPath;
        }

        public static string GetFileName(this ProjectItem projectItem)
        {
            if (projectItem.GetProperty("FileName") is string value)
                return value;
            return string.Empty;
        }

        public static string GetFullPath(this ProjectItem projectItem)
        {
            if (projectItem.GetProperty("FullPath") is string value)
                return value;
            return string.Empty;
        }

        public static bool IsEmbeddedResource(this ProjectItem projectItem)
        {
            return projectItem.GetItemType() == EmbeddedResourceType && projectItem.GetCustomTool() == CustomTool;
        }

        private static object GetProperty(this ProjectItem projectItem, string propertyName)
        {
            if (projectItem.Properties != null)
            {
                for (var i = 1; i <= projectItem.Properties.Count; i++)
                {
                    var item = projectItem.Properties.Item(i);
                    if (item.Name == propertyName)
                    {
                        return item.Value;
                    }
                }
            }
            return null;
        }

        public static void PrintProperties(this ProjectItem projectItem)
        {
            Trace.WriteLine(projectItem.Name);

            if (projectItem.Properties != null)
            {
                for (var i = 1; i <= projectItem.Properties.Count; i++)
                {
                    var item = projectItem.Properties.Item(i);
                    Trace.WriteLine($"\t{item.Name}: {item.Value}");
                }
            }
        }

        public static CultureInfo GetResourceCulture(this ProjectItem projectItem)
        {
            var match = Regex.Match(projectItem.Name, "(.+)[.]([^.]+)([.]resx)$");
            if (match.Success == true)
            {
                return CultureInfo.GetCultureInfo(match.Groups[2].Value);
            }
            return null;
        }

        public static void Write(this ProjectItem projectItem, CremaDataTable projectInfoTable)
        {
            if (FindTable(projectItem, projectInfoTable) is CremaDataTable dataTable)
            {
                var cultureInfo = projectItem.GetResourceCulture();
                var valueName = cultureInfo == null ? "Value" : cultureInfo.Name.Replace('-', '_');
                var path = projectItem.GetFullPath();
                using (var writer = new ResXResourceWriter(path))
                {
                    foreach (var item in dataTable.Rows)
                    {
                        var name = $"{item["Type"]}" == "None" ? $"{item["Name"]}" : $"{item["Type"]}_{item["Name"]}";
                        var node = new ResXDataNode(name, item[valueName])
                        {
                            Comment = $"{item["Comment"]}",
                        };
                        writer.AddResource(node);
                    }
                    writer.Close();
                }

                if (projectItem.GetCustomTool() != string.Empty)
                {
                    WriteDesigner(projectItem, dataTable);
                }
            }
            else
            {
                throw new Exception("항목을 찾을 수 없습니다.");
            }
        }

        private static void WriteDesigner(this ProjectItem projectItem, CremaDataTable dataTable)
        {
            var project = projectItem.ContainingProject;
            var projectPath = Path.GetDirectoryName(project.GetFullName());
            var designerFileName = Path.Combine(projectPath, projectItem.GetCustomToolOutput());
            var resxFileName = projectItem.GetFullPath();
            var ss = StringUtility.SplitPath(Path.GetDirectoryName(projectItem.GetLocalPath()));
            var codeNamespace = $"{project.GetRootNamespace()}.{string.Join(".", ss)}";
            var baseName = Path.GetFileNameWithoutExtension(project.GetLocalPath());
            var isPublic = projectItem.IsPublicResource();

            using (var sw = new StreamWriter(designerFileName))
            {
                var errors = null as string[];
                var provider = new CSharpCodeProvider();
                var code = StronglyTypedResourceBuilder.Create(resxFileName, baseName, codeNamespace, provider, isPublic == false, out errors);
                if (errors.Length > 0)
                {
                    foreach (var error in errors)
                    {
                        Console.WriteLine(error);
                    }
                    return;
                }

                provider.GenerateCodeFromCompileUnit(code, sw, new CodeGeneratorOptions());
                Console.WriteLine(designerFileName);
            }
        }

        private static bool IsPublicResource(this ProjectItem projectItem)
        {
            var project = projectItem.ContainingProject;
            using (var reader = XmlReader.Create(project.GetFullName()))
            {
                var doc = XDocument.Load(reader);
                var ns = doc.Root.GetDefaultNamespace();
                var namespaceManager = new XmlNamespaceManager(reader.NameTable);
                namespaceManager.AddNamespace("xs", ns.NamespaceName);

                var elements = doc.Root.XPathSelectElements($"/xs:Project/xs:ItemGroup/xs:EmbeddedResource", namespaceManager);

                foreach (var element in elements)
                {
                    var attr = element.Attribute(XName.Get("Include", string.Empty));

                    var match = Regex.Match(attr.Value, "(.+)[.]([^.]+)([.]resx)$");
                    if (match.Success == true)
                    {
                        //this.CultureInfo = CultureInfo.GetCultureInfo(match.Groups[2].Value);
                        //this.Name = match.Groups[1].Value + match.Groups[3].Value;
                        //this.FileName = attr.Value;
                    }
                    else
                    {
                        //this.Name = attr.Value;
                        //this.FileName = attr.Value;
                    }

                    var name = attr.Value.Replace(Path.DirectorySeparatorChar, PathUtility.SeparatorChar);
                    if (name != projectItem.GetLocalPath())
                        continue;


                    var e1 = element.XPathSelectElement("./xs:Generator", namespaceManager);
                    if (e1 != null)
                    {
                        return e1.Value == "PublicResXFileCodeGenerator";
                    }
                }
            }

            return false;
        }

        private static CremaDataTable FindTable(ProjectItem projectItem, CremaDataTable projectInfoTable)
        {
            var project = projectItem.ContainingProject;
            var projectPath = project.GetLocalPath();
            var projectItemPath = projectItem.GetLocalPathWithoutCultureInfo();
            var dataSet = projectInfoTable.DataSet;

            foreach (var item in projectInfoTable.Rows)
            {
                if (item["ProjectPath"] is string text && text == projectPath)
                {
                    if (GetExportInfo(item) is CremaDataRow exportInfoRow)
                    {
                        return dataSet.Tables[exportInfoRow["TableName"] as string];
                    }
                }
            }

            return null;

            CremaDataRow GetExportInfo(CremaDataRow dataRow)
            {
                foreach (var item in dataRow.GetChildRows("ExportInfo"))
                {
                    if (item["FileName"] is string text && text == projectItemPath)
                    {
                        return item;
                    }
                }
                return null;
            }
        }
    }
}
