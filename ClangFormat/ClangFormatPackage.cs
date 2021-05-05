// VsPkg.cs : Implementation of ClangFormat
//

using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.ComponentModel.Design;
using Microsoft.Win32;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using EnvDTE;
using EnvDTE80;
using EnvDTE90;

namespace LLVM.ClangFormat
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the registration utility (regpkg.exe) that this class needs
    // to be registered as package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // A Visual Studio component can be registered under different regitry roots; for instance
    // when you debug your package you want to register it in the experimental hive. This
    // attribute specifies the registry root to use if no one is provided to regpkg.exe with
    // the /root switch.
    [DefaultRegistryRoot("Software\\Microsoft\\VisualStudio\\9.0")]
    // This attribute is used to register the informations needed to show the this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration(false, "#110", "#112", "1.0", IconResourceID = 400)]
    // In order be loaded inside Visual Studio in a machine that has not the VS SDK installed, 
    // package needs to have a valid load key (it can be requested at 
    // http://msdn.microsoft.com/vstudio/extend/). This attributes tells the shell that this 
    // package has a load key embedded in its resources.
    [ProvideLoadKey("Professional", "1.0", "ClangFormat", "LLVM", 1)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource(1000, 1)]
    [Guid(GuidList.guidClangFormatPkgString)]
    public sealed class ClangFormatPackage : Package
    {
        OutputWindowPane owp;
        ClangFormat cf;

        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public ClangFormatPackage()
        {
            Trace.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));
        }

        /////////////////////////////////////////////////////////////////////////////
        // Overriden Package Implementation
        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initilaization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            Trace.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
            base.Initialize();

            // Add our command handlers for menu (commands must exist in the .vsct file)
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                // Create the command for the menu item.
                CommandID menuCommandID = new CommandID(GuidList.guidClangFormatCmdSet, (int)PkgCmdIDList.ClangFormatCommand);
                MenuCommand menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
                mcs.AddCommand(menuItem);
            }

            DTE env = (DTE)GetService(typeof(DTE));
            EnvDTE.Properties props =
                env.get_Properties("LLVM", "clang-format");

            owp = CreatePane("clang-format");
            cf = new ClangFormat(owp, props);
        }
        #endregion

        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            DTE env = (DTE)GetService(typeof(DTE));

            cf.FormatFile(env);
            //// Show a Message Box to prove we were here
            //IVsUIShell uiShell = (IVsUIShell)GetService(typeof(SVsUIShell));
            //Guid clsid = Guid.Empty;
            //int result;
            //Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(uiShell.ShowMessageBox(
            //           0,
            //           ref clsid,
            //           "ClangFormat",
            //           string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.ToString()),
            //           string.Empty,
            //           0,
            //           OLEMSGBUTTON.OLEMSGBUTTON_OK,
            //           OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
            //           OLEMSGICON.OLEMSGICON_INFO,
            //           0,        // false
            //           out result));
        }

        [ClassInterface(ClassInterfaceType.AutoDual)]
        [CLSCompliant(false), ComVisible(true)]
        public class OptionPageGrid : DialogPage
        {
            private string clangExecutable = "";
            private string cfStyle = "file";
            private string fallbackStyle = "none";
            private string assumeFilename = "";
            private string cursorPosition = "Restore";
            private bool saveOnFormat = true;
            private bool outputEnabled = false;
            private bool sortIncludes = false;
            private bool searchClangFormat = true;

            [Category("clang-format settings")]
            [Description("clang-format executable (full path including *.exe)")]
            [DisplayName("clang-format executable")]
            public string ClangExecutable
            {
                get { return clangExecutable; }
                set { clangExecutable = value; }
            }

            [Category("clang-format settings")]
            [Description("Coding style")]
            [DisplayName("Coding style")]
            public string CfStyle
            {
                get { return cfStyle; }
                set { cfStyle = value; }
            }

            [Category("clang-format settings")]
            [Description("Use fallback style")]
            [DisplayName("Fallback style")]
            public string FallbackStyle
            {
                get { return fallbackStyle; }
                set { fallbackStyle = value; }
            }

            [Category("clang-format settings")]
            [Description("Assume filename")]
            [DisplayName("Assume filename")]
            public string AssumeFilename
            {
                get { return assumeFilename; }
                set { assumeFilename = value; }
            }

            [Category("clang-format settings")]
            [Description("Sort includes (overwrites option defined in .clang-format)")]
            [DisplayName("Sort includes")]
            public bool SortIncludes
            {
                get { return sortIncludes; }
                set { sortIncludes = value; }
            }

            [Category("clang-format settings")]
            [Description("Search clang-format in registry. Search happens on first execution due to lazy loading.")]
            [DisplayName("Search clang-format")]
            public bool SearchClangFormat
            {
                get { return searchClangFormat; }
                set { searchClangFormat = value; }
            }

            [Category("clang-format settings")]
            [Description("Places the cursor at the given position. Possible values are: Restore, Top, SameLine, Bottom")]
            [DisplayName("Cursor insertion point")]
            public string CursorPosition
            {
                get { return cursorPosition; }
                set { cursorPosition = value; }
            }

            [Category("clang-format settings")]
            [Description("Saves the document after formatting")]
            [DisplayName("Save on format")]
            public bool SaveOnFormat
            {
                get { return saveOnFormat; }
                set { saveOnFormat = value; }
            }

            [Category("clang-format settings")]
            [Description("Enable/Disable output messages.")]
            [DisplayName("Enable output")]
            public bool OutputEnabled
            {
                get { return outputEnabled; }
                set { outputEnabled = value; }
            }
        }

        [ProvideOptionPage(typeof(OptionPageGrid), "LLVM", "clang-format", 0, 0, true)]
        public sealed class ClangFormatCSPackage : Package
        {
        }

        OutputWindowPane CreatePane(string title)
        {
            DTE2 dte = (DTE2)GetService(typeof(DTE));
            OutputWindowPanes panes =
                dte.ToolWindows.OutputWindow.OutputWindowPanes;

            try
            {
                // If the pane exists already, return it.
                return panes.Item(title);
            }
            catch (ArgumentException)
            {
                // Create a new pane.
                return panes.Add(title);
            }
        }
    }
}