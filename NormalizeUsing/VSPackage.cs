using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using EnvDTE;
using EnvDTE80;

namespace NormalizeUsing
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [ProvideAutoLoad(UIContextGuids80.NoSolution)]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists)]
    [ProvideAutoLoad(UIContextGuids80.SolutionHasSingleProject)]
    [ProvideAutoLoad(UIContextGuids80.SolutionHasMultipleProjects)]
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(VSPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class VSPackage : Package
    {
        /// <summary>
        /// VSPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "b8ef66b0-63c5-4bb8-ab1b-1cdb579ec385";

        /// <summary>
        /// Initializes a new instance of the <see cref="VSPackage"/> class.
        /// </summary>
        public VSPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        #region Package Members

        private DocumentEvents documentEvents;

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            base.Initialize();

            var dte = GetService(typeof(DTE)) as DTE2;
            documentEvents = dte.Events.DocumentEvents;
            documentEvents.DocumentSaved += DocumentEvents_DocumentSaved;
        }

        void DocumentEvents_DocumentSaved(Document document)
        {
            var path = document.FullName;
            var lang = document.Language;
            if (document.Kind != "{8E7B96A8-E33D-11D0-A6D5-00C04FB67F6A}")
            {
                // then it's not a text file
                return;
            }
            if (!lang.Equals("CSharp"))
            {
                return;
            }

            var regEx = new Regex("(using\\s+((\\w+)(\\.(\\w+))*);)", RegexOptions.None);
            var firstList = new List<string>();
            var secondList = new List<string>();
            var usingList = new List<string>();
            var isFound = false;
 
            using (var stream = new FileStream(path, FileMode.Open))
            {
                using (var reader = new StreamReader(stream, Encoding.Default, true))
                {
                    while(true)
                    {
                        var text = reader.ReadLine();
                        if (text == null)
                        {
                            usingList.Sort();
                            break;
                        }

                        var match = regEx.Match(text);
                        if (match.Success)
                        {
                            var value = match.Groups[1].Value;
                            usingList.Add(value);
                            isFound = true;
                        }
                        else if (isFound)
                        {
                            secondList.Add(text);
                        }
                        else
                        {
                            firstList.Add(text);
                        }
                    }
                }
            }

            var temp_path = document.FullName + ".tmp";
            using (var stream = new FileStream(temp_path, FileMode.Create))
            {
                using (var writer = new StreamWriter(stream, Encoding.Default))
                {
                    foreach (var line in firstList)
                    {
                        writer.Write(line + Environment.NewLine);
                    }
                    foreach (var line in usingList)
                    {
                        writer.Write(line + Environment.NewLine);
                    }
                    foreach (var line in secondList)
                    {
                        writer.Write(line + Environment.NewLine);
                    }
                }
            }

            File.Delete(path);
            File.Move(temp_path, path);

            File.SetLastWriteTime(path, DateTime.Now);
        }
        #endregion
    }
}
