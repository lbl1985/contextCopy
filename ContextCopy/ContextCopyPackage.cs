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
        private void MenuItemCallback(object sender, EventArgs e)
        {
            var dte2 = (EnvDTE80.DTE2)Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(SDTE));
            //IVsWindowFrame wf = Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(SVsWindowFrame)) as IVsWindowFrame;
            //IVsTextView view = this.ServiceProvider.GetService(typeof(SVsOutputWindow)) as IVsTextView;
            //IVsWindowFrame wf = this.ServiceProvider.GetService(typeof(SVsWindowFrame)) as IVsWindowFrame;
            //IntPtr pDBM;
            //Guid riid = typeof(IVsDropdownBarManager).GUID;
            //wf.QueryViewInterface(ref riid, out pDBM);
            //IVsDropdownBarManager dropdownBarManager = (IVsDropdownBarManager)Marshal.GetObjectForIUnknown(pDBM);

            //var txtMgr = (IVsTextManager)ServiceProvider.GetService(typeof(SVsTextManager));
            //txtMgr.
            //IVsDropdownBarClient dbc = (IVsDropdownBarClient)ServiceProvider.GetService(typeof(IVsDropdownBarClient));
            //IVsCodeWindow codeWindow = this.ServiceProvider.GetService(typeof(IVsCodeWindow));
            //codeWindow.service

            // http://stackoverflow.com/questions/28321106/how-can-i-get-an-ivsdropdownbar-out-of-an-envdte-window
            //IVsCodeWindow codeWindow = (IVsCodeWindow)Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(SVsCodeWindow));
            
            //IVsCodeWindow codeWindow = (IVsCodeWindow)this.GetService(typeof(SVsCodeWindow));
            //dte2.ActiveWindow.
            
            ////IVsDropdownBarClient dbc;
            //string funName;
            ////ddb.GetClient(out dbc);
            //dbc.GetEntryText(1, 0, out funName);

            //IVsTextView textView;
            //txtMgr.GetActiveView(1, null, out textView);
            //IVsEnumGUID currentLanGUID;
            //txtMgr.EnumLanguageServices(out currentLanGUID);

            //IWpfTextView wpfViewCurrent = AdaptersFactory.GetWpfTextView(textView);
            //ITextBuffer textCurrent = wpfViewCurrent.TextBuffer;
            
            try // When there is no Document opened, do nothing
            {
                Document objD = (Document)dte2.ActiveDocument.Object("Document");
                if (null != objD)
                {
                    // Basic functionality section only file name and line number info
                    var fn = objD.Name;
                    TextDocument objTD = (TextDocument)dte2.ActiveDocument.Object("TextDocument");
                    TextPoint objTP = objTD.Selection.ActivePoint;
                    var sel = objTD.Selection;
                    var cl = objTP.Line;
                    string text = sel.Text;
                    string basicText = text + " [" + fn + ": " + cl + " ";
                                        
                    string funName = "";
                    string className = "";                   
                    
                    EnvDTE.CodeClass codeClass = objTP.CodeElement[vsCMElement.vsCMElementClass] as EnvDTE.CodeClass;
                    if (codeClass != null)
                    {
                        className = codeClass.Name;
                        basicText += className + "::";
                    }

                    EnvDTE.CodeFunction codeFun = objTP.CodeElement[vsCMElement.vsCMElementFunction] as EnvDTE.CodeFunction;
                    if (codeFun != null)
                    {
                        funName = codeFun.get_Prototype(8);
                        basicText += funName;
                    }

                    basicText += "]";
                    Clipboard.SetText(basicText);

                    //codeFun.Prototype[vsCMPrototype.vsCMPrototypeClassName];
                    //var className = codeFun.get_Prototype([vsCMPrototypeClassName]);

                    //EnvDTE.CodeClass codeClass = objTP.CodeElement[vsCMElement.vsCMElementClass] as EnvDTE.CodeClass;

                    //if (codeClass != null && codeFun != null)
                    //{

                    //}

                    //var funName = codeFun.get_Prototype();
                    
                    //foreach (CodeParameter param in func.Parameters)
                    //{
                    //    TextPoint start = param.GetStartPoint(vsCMPart.vsCMPartWhole);
                    //    TextPoint finish = param.GetEndPoint(vsCMPart.vsCMPartWhole);
                    //    parms += start.CreateEditPoint().GetText(finish);
                    //}
                    //string funName = ce.FullName;
                    //var cl = sel.CurrentLine;
                    

                    //vsCMElement scopes = 0;
                    //foreach(vsCMElement scope in Enum.GetValues(scopes.GetType()))
                    //{
                    //    ;
                    //}
                    //dte2.ActiveDocument.Collection.
                    //ProjectItem objPI = dte2.ActiveDocument.ProjectItem;
                    //string className = dte2.ActiveDocument.ProjectItem.FileCodeModel.CodeElementFromPoint(objTP, vsCMElement.vsCMElementModule).ToString();
                    //string funName = dte2.ActiveDocument.ProjectItem.FileCodeModel.CodeElementFromPoint(objTP, EnvDTE.vsCMElement.vsCMElementFunction).ToString();
                    //string funName = dte2.ActiveDocument.ProjectItem.FileCodeModel.CodeElementFromPoint(objTP, vsCMElement.vsCMElementFunctionInvokeStmt).ToString();
                    //dte2.ActiveDocument.
                    //string finalText = text + " [" + fn + ": " + cl + ", "  + "]";
                    //Clipboard.SetText(finalText);
                }
            }
            catch
            {
            }
        }

    }
}
