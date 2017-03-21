using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace MakeURL4PPT
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
            return GetResourceText("MakeURL4PPT.Ribbon.xml");
        }

        #endregion

        #region リボンのコールバック

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void MakeURL(Office.IRibbonControl control)
        {
            PowerPoint.Application appl = Globals.ThisAddIn.Application;
            PowerPoint.Presentation thePresentation = appl.ActivePresentation;
            PowerPoint.Slide theSlide = appl.ActiveWindow.View.Slide;
            PowerPoint.Selection selection = appl.ActiveWindow.Selection;
            String urlstring = UrlHandler.Program.PREFIX + thePresentation.FullName;
            urlstring += "#" + theSlide.SlideIndex;
            if (control.Id.StartsWith("MakeURLShape") ||
                control.Id.StartsWith("MakeURLTextEdit") ||
                control.Id.StartsWith("MakeURLObjectsGroup"))
            {
                urlstring += "!";
                String[] names = new String[selection.ShapeRange.Count];
                for (int i = 0; i < selection.ShapeRange.Count; i++)
                {
                   urlstring += selection.ShapeRange[1+i].Name + ",";
                }
                urlstring = urlstring.Substring(0, urlstring.Length - 1); // trim trailing [!,]
            }
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
