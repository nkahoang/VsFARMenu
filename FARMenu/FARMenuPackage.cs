using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using EnvDTE;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.Win32;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;

namespace nkaHnt.FARMenu
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
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("FARMenu", "Find and Replace Extended Menu", "1.0", IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidFARMenuPkgString)]
    public sealed class FARMenuPackage : Package
    {
        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public FARMenuPackage()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));
        }



        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation
        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            Debug.WriteLine (string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
            base.Initialize();

            // Add our command handlers for menu (commands must exist in the .vsct file)
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if ( null != mcs )
            {
                // Create the command for the menu item.
                var FindInMenuId = new CommandID(GuidList.guidFARMenuCmdSet, (int)PkgCmdIDList.nkahntFindIn);
                var ReplaceInMenuId = new CommandID(GuidList.guidFARMenuCmdSet, (int)PkgCmdIDList.nkahntReplaceIn);

                var findMenuItem = new OleMenuCommand(FindMenuItemCallback, FindInMenuId);
                findMenuItem.BeforeQueryStatus += menuItem_BeforeQueryStatus;

                var replaceMenuItem = new OleMenuCommand(ReplaceMenuItemCallback, ReplaceInMenuId);
                replaceMenuItem.BeforeQueryStatus += menuItem_BeforeQueryStatus;

                mcs.AddCommand(findMenuItem);
                mcs.AddCommand(replaceMenuItem);
            }
        }

        string getFindTarget(IVsHierarchy hierarchy, uint itemId)
        {
            // Get the file/directory path
            string itemFullPath = null;

            var vsProject = hierarchy as IVsProject;

            if (vsProject == null && itemId == VSConstants.VSITEMID_ROOT)
            {
                var solution = Package.GetGlobalService(typeof(SVsSolution)) as IVsSolution;
                string solutionFile = null;
                string solutionUserFile = null;
                solution.GetSolutionInfo(out itemFullPath, out solutionFile, out solutionUserFile);

            }
            else if (itemId == VSConstants.VSITEMID_ROOT)
            {
                string projectFullPath = null;
                //this is selecting a project, not an item
                vsProject.GetMkDocument(VSConstants.VSITEMID_ROOT, out projectFullPath);
                itemFullPath = Path.GetDirectoryName(projectFullPath);
            }
            else
            {
                vsProject.GetMkDocument(itemId, out itemFullPath);
            }

            if (!Directory.Exists(itemFullPath) &&
                !File.Exists(itemFullPath))
            {
                return null;
            }
            
            return itemFullPath;
        }

        void menuItem_BeforeQueryStatus(object sender, EventArgs e)
        {

            // get the menu that fired the event
            var menuCommand = sender as OleMenuCommand;
            if (menuCommand != null)
            {
                // start by assuming that the menu will not be shown
                menuCommand.Visible = false;
                menuCommand.Enabled = false;

                IVsHierarchy hierarchy = null;
                uint itemid = VSConstants.VSITEMID_NIL;
                
                if (!IsSingleProjectItemSelection(out hierarchy, out itemid)) return;
                
                if (string.IsNullOrEmpty(getFindTarget(hierarchy, itemid)))
                {
                    return;
                }
                
                menuCommand.Visible = true;
                menuCommand.Enabled = true;
            }
        }


        public static bool IsSingleProjectItemSelection(out IVsHierarchy hierarchy, out uint itemid)
        {
            hierarchy = null;
            itemid = VSConstants.VSITEMID_NIL;
            int hr = VSConstants.S_OK;

            var monitorSelection = Package.GetGlobalService(typeof(SVsShellMonitorSelection)) as IVsMonitorSelection;
            var solution = Package.GetGlobalService(typeof(SVsSolution)) as IVsSolution;
            if (monitorSelection == null || solution == null)
            {
                return false;
            }

            IVsMultiItemSelect multiItemSelect = null;
            IntPtr hierarchyPtr = IntPtr.Zero;
            IntPtr selectionContainerPtr = IntPtr.Zero;

            try
            {
                hr = monitorSelection.GetCurrentSelection(out hierarchyPtr, out itemid, out multiItemSelect, out selectionContainerPtr);

                if (hierarchyPtr == IntPtr.Zero && itemid == VSConstants.VSITEMID_ROOT) //solution selected
                {
                    return true;
                }

                if (ErrorHandler.Failed(hr) || hierarchyPtr == IntPtr.Zero || itemid == VSConstants.VSITEMID_NIL)
                {
                    // there is no selection
                    return false;
                }

                // multiple items are selected
                if (multiItemSelect != null) return false;

                // there is a hierarchy root node selected, thus it is not a single item inside a project

                //if (itemid == VSConstants.VSITEMID_ROOT) return false;

                hierarchy = Marshal.GetObjectForIUnknown(hierarchyPtr) as IVsHierarchy;
                if (hierarchy == null) return false;

                Guid guidProjectID = Guid.Empty;

                if (ErrorHandler.Failed(solution.GetGuidOfProject(hierarchy, out guidProjectID)))
                {
                    return false; // hierarchy is not a project inside the Solution if it does not have a ProjectID Guid
                }

                // if we got this far then there is a single project item selected
                return true;
            }
            finally
            {
                if (selectionContainerPtr != IntPtr.Zero)
                {
                    Marshal.Release(selectionContainerPtr);
                }

                if (hierarchyPtr != IntPtr.Zero)
                {
                    Marshal.Release(hierarchyPtr);
                }
            }
        }
        #endregion

        private void FindMenuItemCallback(object sender, EventArgs e)
        {
            MenuItemCallback(sender, e, "Edit.FindInFiles");
        }
        private void ReplaceMenuItemCallback(object sender, EventArgs e)
        {
            MenuItemCallback(sender, e, "Edit.ReplaceInFiles");
        }
        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void MenuItemCallback(object sender, EventArgs e, string command)
        {

            IVsHierarchy hierarchy = null;
            uint itemid = VSConstants.VSITEMID_NIL;

            if (!IsSingleProjectItemSelection(out hierarchy, out itemid)) return;
            // Get the file path
            string itemFullPath = getFindTarget(hierarchy, itemid);

            if (string.IsNullOrWhiteSpace(itemFullPath))
            {
                return;
            }

            var dte = this.GetService(typeof(SDTE)) as EnvDTE80.DTE2;
            if (dte == null)
                return;
            var findWindow = dte.Find;

            findWindow.SearchPath = itemFullPath;
            findWindow.Action = vsFindAction.vsFindActionFindAll;
            dte.ExecuteCommand(command);
        }

    }
}
