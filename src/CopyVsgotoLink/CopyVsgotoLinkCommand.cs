using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;
using Task = System.Threading.Tasks.Task;

namespace CopyVsgotoLink {
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class CopyVsgotoLinkCommand {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("322445b2-953d-46ad-bd94-b101b10bec3e");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="CopyVsgotoLinkCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private CopyVsgotoLinkCommand(AsyncPackage package, OleMenuCommandService commandService) {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static CopyVsgotoLinkCommand Instance {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider {
            get {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package) {
            // Switch to the main thread - the call to AddCommand in CopyVsgotoLinkCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new CopyVsgotoLinkCommand(package, commandService);
        }

        internal static IVsTextView GetIVsTextView(out string filePath) {
            DTE2 dte = (DTE2)Package.GetGlobalService(typeof(SDTE));
            filePath = dte.ActiveDocument.Path + dte.ActiveDocument.Name;
            IVsUIHierarchy vsUIHierarchy;
            uint num;
            IVsWindowFrame vsWindowFrame;
            if (VsShellUtilities.IsDocumentOpen(new ServiceProvider((Microsoft.VisualStudio.OLE.Interop.IServiceProvider)dte), filePath, Guid.Empty, out vsUIHierarchy, out num, out vsWindowFrame)) {
                return VsShellUtilities.GetTextView(vsWindowFrame);
            }
            return null;
        }

        internal static string GetUrl() {
            string filePath = "";
            int CaretLine = 0;
            int CaretCol = 0;
            IVsTextView ivsTextView = GetIVsTextView(out filePath);
            ivsTextView.GetCaretPos(out CaretLine, out CaretCol);

            var replace = new KeyValuePair<string, string>("", "");

            foreach (DictionaryEntry env in Environment.GetEnvironmentVariables()) {
                var tmp = new KeyValuePair<string, string>(env.Key.ToString(), env.Value.ToString());

                if(filePath.StartsWith(tmp.Value, true, CultureInfo.InvariantCulture)) {
                    if(tmp.Value.Length > replace.Value.Length) {
                        replace = tmp;
                    }
                }
            }

            if(replace.Key.Length > 0) {
                filePath = "{" + replace.Key + "}" + filePath.Substring(replace.Value.Length);
            }

            return $"vsgoto:{filePath}:{CaretLine + 1}";
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e) {
            ThreadHelper.ThrowIfNotOnUIThread();
            string message = GetUrl();
            string title = "CopyVsgotoLink";

            Clipboard.SetText(message);

            // Show a message box to prove we were here
            VsShellUtilities.ShowMessageBox(
                this.package,
                message,
                title,
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}
