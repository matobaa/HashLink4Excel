using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace MakeURL4XLS
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility のメンバー

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("MakeURL4XLS.Ribbon.xml");
        }

        #endregion

        #region リボンのコールバック

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void MakeURL(Office.IRibbonControl control)
        {
            Excel.Application appl = Globals.ThisAddIn.Application;
            Excel.Workbook theWorkbook = appl.ActiveWorkbook;
            Excel.Worksheet theSheet = theWorkbook.ActiveSheet;
            Excel.Range selection = appl.Selection;
            String urlstring = UrlHandler.Program.PREFIX + theWorkbook.FullName;
            urlstring += "#" + theSheet.Name;
            if (control.Id.StartsWith("MakeURLCell") ||
                control.Id.StartsWith("MakeURLRow") ||
                control.Id.StartsWith("MakeURLColumn"))
                urlstring += "!" + selection.Address;
            // paste text to clipboard
            DataObject data = new DataObject();
            data.SetData(DataFormats.Text, urlstring);
            // paste html to clipboard
            data.SetData(DataFormats.Html, HTMLClipboardFormat(urlstring));
            Clipboard.SetDataObject(data, true);
        }

        public string GetLabel(Office.IRibbonControl Control)
        {
            return "ここへのハイパーリンクをコピー";
        }

        private static string HTMLClipboardFormat(string urlstring)
        {
            System.Reflection.Assembly asm;
            asm = System.Reflection.Assembly.GetExecutingAssembly();
            System.Resources.ResourceManager rm =
                new System.Resources.ResourceManager(
                asm.GetName().Name + ".Properties.Resources", asm);
            string s = rm.GetString("HTMLClipboardFormat");
            return String.Format(s, urlstring, 167 + urlstring.Length * 2, 134 + urlstring.Length * 2);
        }

        #endregion

        #region ヘルパー

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
