using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using Microsoft.Win32;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using EnvDTE80;
using EnvDTE;
using System.Windows.Forms;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Package;
using Microsoft.VisualStudio.TextManager.Interop;



namespace Dragonist.ContextCopy
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
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidContextCopyPkgString)]
    public sealed class ContextCopyPackage : Package
    {
        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public ContextCopyPackage()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));
            this.m_iMode = 0;
            this.m_bFirstHit = true;
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
                CommandID menuCommandID = new CommandID(GuidList.guidContextCopyCmdSet, (int)PkgCmdIDList.cmdidContextCopy);
                MenuCommand menuItem = new MenuCommand(MenuItemCallback, menuCommandID );
                mcs.AddCommand( menuItem );
            }
        }
        #endregion

        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        /// 

        private void updateStatusBar(IVsStatusbar bar, string status)
        {
            int freeze;
            bar.IsFrozen(out freeze);

            if (freeze != 0)
            {
                bar.FreezeOutput(0);
            }
            
            bar.SetText(status);
            bar.FreezeOutput(1);
        }

        // Mode 0: copied text [<filepath + filename>: <line#> <className>::<function signature>]
        // Mode 1: copied text [<filename>:<line#> <className>::<function signature>]
        // Mode 2: copied text [<filename>:<line#> <className>::<function name>]
        // Mode 3: copied text [<filename>:<line#> <function name>]
        // Mode 4: copied text [<filename>:<line#>]
        private void getContext(ref EnvDTE80.DTE2 dte2, ref Document objD, IVsStatusbar bar, int mode)
        {
            // if no document opened, return "";
            if(objD == null)
            {
                return;
            }

            string fn = "file:///" + objD.Path + objD.Name + " ";
            if (mode != 0)
            {
                fn = objD.Name;
            } 
            
            TextDocument objTD = (TextDocument)dte2.ActiveDocument.Object("TextDocument");
            TextPoint objTP = objTD.Selection.ActivePoint;
            var sel = objTD.Selection;
            var cl = objTP.Line;
            string text = sel.Text;
            string basicText = text + " [" + fn + ":" + cl + " ";

            string funName = "";
            string className = "";

            EnvDTE.CodeClass codeClass = objTP.CodeElement[vsCMElement.vsCMElementClass] as EnvDTE.CodeClass;
            if (codeClass != null && mode < 3)
            {
                className = codeClass.Name;
                basicText += className + "::";
            }

            EnvDTE.CodeFunction codeFun = objTP.CodeElement[vsCMElement.vsCMElementFunction] as EnvDTE.CodeFunction;
            if (codeFun != null)
            {
                if(mode == 0 || mode == 1)
                {
                    funName = codeFun.get_Prototype(8);
                } else if (mode == 2 || mode == 3)
                {
                    funName = codeFun.Name;
                }                
                basicText += funName;
            }

            basicText += "]";

            updateStatusBar(bar, "Mode:" + this.m_iMode.ToString() + " " + basicText);
            Clipboard.SetText(basicText);
        }

        private void updateTimeMode()
        {
            if (this.m_bFirstHit)
            {
                this.m_dtLastHit = System.DateTime.UtcNow;
                this.m_bFirstHit = false;
            }
            else
            {
                DateTime now = System.DateTime.UtcNow;
                var deltaSecond = now.Subtract(this.m_dtLastHit).TotalMilliseconds / 1000;

                if (deltaSecond < 2.5)
                {
                    this.m_iMode = (this.m_iMode + 1) % 5;
                }

                this.m_dtLastHit = now;
            }
        }
        private void MenuItemCallback(object sender, EventArgs e)
        {
            updateTimeMode();

            var dte2 = (EnvDTE80.DTE2)Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(SDTE));
            IVsStatusbar statusBar = (IVsStatusbar)GetService(typeof(SVsStatusbar));

            try // When there is no Document opened, do nothing
            {
                Document objD = (Document)dte2.ActiveDocument.Object("Document");
                if (null != objD)
                {
                    // get Context, based on mode to get differnt context info
                    getContext(ref dte2, ref objD, statusBar, this.m_iMode);
                }
            }
            catch
            {
            }
        }
        private int m_iMode;
        private DateTime m_dtLastHit;
        private bool m_bFirstHit;

    }
}
