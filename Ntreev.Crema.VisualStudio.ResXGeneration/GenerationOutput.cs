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
    static class GenerationOutput
    {
        private static OutputWindowPane outputWindowPane;

        public static void Write(string text)
        {
            Pane.OutputString(text);
        }

        public static void WriteLine()
        {
            WriteLine(string.Empty);
        }

        public static void WriteLine(string text)
        {
            Pane.OutputString(text + Environment.NewLine);
        }

        private static OutputWindowPane Pane
        {
            get
            {
                if (outputWindowPane == null)
                {
                    var dte = GenerationCommand.Instance.ServiceProvider.GetService(typeof(SDTE)) as EnvDTE80.DTE2;

                    var outputPanes = dte.ToolWindows.OutputWindow.OutputWindowPanes;

                    var enumerator = outputPanes.GetEnumerator();
                    while (enumerator.MoveNext())
                    {
                        if (enumerator is OutputWindowPane pane && pane.Name == "Crema Resources")
                        {
                            outputWindowPane = pane;
                            break;
                        }
                    }

                    if (outputWindowPane == null)
                    {
                        outputWindowPane = outputPanes.Add("Crema Resources");
                    }
                }
                return outputWindowPane;
            }
        }
    }
}
