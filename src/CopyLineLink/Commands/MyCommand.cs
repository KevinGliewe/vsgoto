using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;
using System.Collections.Generic;
using System.Collections;
using System.Globalization;
using EnvDTE80;
using System.Windows.Forms;

namespace CopyLineLink
{
    [Command(PackageIds.MyCommand)]
    internal sealed class MyCommand : BaseCommand<MyCommand>
    {
        internal IVsTextView GetIVsTextView(out string filePath)
        {
            DTE2 dte = (DTE2)Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(SDTE));
            filePath = dte.ActiveDocument.Path + dte.ActiveDocument.Name;
            IVsUIHierarchy vsUIHierarchy;
            uint num;
            IVsWindowFrame vsWindowFrame;
            if (VsShellUtilities.IsDocumentOpen(new ServiceProvider((Microsoft.VisualStudio.OLE.Interop.IServiceProvider)dte), filePath, Guid.Empty, out vsUIHierarchy, out num, out vsWindowFrame))
            {
                return VsShellUtilities.GetTextView(vsWindowFrame);
            }
            return null;
        }

        internal string GetUrl()
        {
            string filePath = "";
            int CaretLine = 0;
            int CaretCol = 0;
            IVsTextView ivsTextView = GetIVsTextView(out filePath);
            ivsTextView.GetCaretPos(out CaretLine, out CaretCol);

            var replace = new KeyValuePair<string, string>("", "");

            foreach (DictionaryEntry env in Environment.GetEnvironmentVariables())
            {
                var tmp = new KeyValuePair<string, string>(env.Key.ToString(), env.Value.ToString());

                if (filePath.StartsWith(tmp.Value, true, CultureInfo.InvariantCulture))
                {
                    if (tmp.Value.Length > replace.Value.Length)
                    {
                        replace = tmp;
                    }
                }
            }

            if (replace.Key.Length > 0)
            {
                filePath = "{" + replace.Key + "}" + filePath.Substring(replace.Value.Length);
            }

            return $"vsgoto:{filePath}:{CaretLine + 1}";
        }

        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            string message = GetUrl();
            string title = "CopyLineLink";

            Clipboard.SetText(message);

            //// Show a message box to prove we were here
            //VsShellUtilities.ShowMessageBox(
            //    this.package,
            //    message,
            //    title,
            //    OLEMSGICON.OLEMSGICON_INFO,
            //    OLEMSGBUTTON.OLEMSGBUTTON_OK,
            //    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

            await VS.MessageBox.ShowAsync(title, message);
        }
    }
}
