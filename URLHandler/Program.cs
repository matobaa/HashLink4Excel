using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

using Forms = System.Windows.Forms;
using System.Runtime.CompilerServices;
using Microsoft.Office.Core;
using System;

[assembly: InternalsVisibleTo("UnitTest")]

[assembly: InternalsVisibleTo("MakeURL4PPT")]

[assembly: InternalsVisibleTo("MakeURL4XLS")]

namespace UrlHandler
{
    internal class Program
    {
        internal const string PREFIX = "ehl:";
        internal const string TITLE = "URLHandler";

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(int hWnd);

        internal static Excel._Application GetExcel()
        {
            try
            {
                return (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch (COMException)
            {
                return new Excel.Application();
            }
        }

        internal static PowerPoint._Application GetPowerPoint()
        {
            try
            {
                return (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
            }
            catch (COMException)
            {
                return new PowerPoint.Application();
            }
        }

        internal static void Main(string[] args)
        {
            string arg = string.Join(" ", args);
            if (arg.Length < PREFIX.Length)
            {
                Forms.MessageBox.Show("Missing Argument. exitting...", TITLE, Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Exclamation);
                return;
            }

            arg = arg.Substring(PREFIX.Length); // trim a prefix
            args = arg.Split('#');
            string path = args[0];

#if !DEBUG
            // Dialog
            System.Reflection.Assembly asm;
            asm = System.Reflection.Assembly.GetExecutingAssembly();
            System.Resources.ResourceManager rm =
                new System.Resources.ResourceManager(
                asm.GetName().Name + ".Properties.Resources", asm);

            string guardDialogMessage = rm.GetString("GuardDialog");
#endif

            try
            {
                if (path.EndsWith(".xlsx"))
                {
#if !DEBUG
                    Forms.DialogResult result = Forms.MessageBox.Show(
                        String.Format(guardDialogMessage, "Microsoft Excel", arg),
                        TITLE, Forms.MessageBoxButtons.YesNo,
                        Forms.MessageBoxIcon.Exclamation,
                        Forms.MessageBoxDefaultButton.Button2);
                    if (result == Forms.DialogResult.No) return;
#endif
                    Excel._Application appl = GetExcel();
                    Excel.Workbooks workbooks = appl.Workbooks;
                    Excel.Workbook workbook;
                    try // to open path. if fail once, try to open again with URLdecoded path
                    {
                        workbook = workbooks.Open(path);
                    }
                    catch (COMException)
                    {
                        path = Uri.UnescapeDataString(path);
                        workbook = workbooks.Open(path);
                    }

                    appl.Visible = true;
                    if (args.Length > 0) // if fragment exists
                        if (Exists(appl.Names, args[1]))
                            appl.Goto(args[1]);
                        else
                            SelectFragment(workbook, args[1]);
                    // bring up
                    workbook.Activate();
                    SetForegroundWindow(appl.Hwnd);
                }
                else if (path.EndsWith(".pptx"))
                {
#if !DEBUG
                    Forms.DialogResult result = Forms.MessageBox.Show(
                        String.Format(guardDialogMessage, "Microsoft PowerPoint", arg),
                        TITLE, Forms.MessageBoxButtons.YesNo,
                        Forms.MessageBoxIcon.Exclamation,
                        Forms.MessageBoxDefaultButton.Button2);
                    if (result == Forms.DialogResult.No) return;
#endif
                    PowerPoint._Application appl = GetPowerPoint();
                    PowerPoint.Presentations ppts = appl.Presentations;
                    PowerPoint.Presentation ppt;
                    try  // to open path. if fail once, try to open again with URLdecoded path
                    {
                        ppt = ppts.Open(path);
                    }
                    catch (COMException)
                    {
                        ppt = ppts.Open(Uri.UnescapeDataString(path));

                    }
                    if (args.Length > 0) // if fragment exists
                        SelectFragment(ppt, args[1]);
                    // bring up
                    appl.Activate();
                    appl.Visible = MsoTriState.msoTrue;
                    SetForegroundWindow(appl.HWND);
                }
                else
                { // TODO: add another suffixes supporting
                    Forms.MessageBox.Show("only .xlsx is supported. exitting.", TITLE, Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Exclamation);
                    return;
                }
            }
            catch (COMException ex)
            {
                Forms.MessageBox.Show(
                    (GetExcel().Name.Equals(ex.Source) ? ex.Message : "ファイルを開けませんでした。")
#if DEBUG
                    + "\r\n\r\n" + ex.StackTrace
#endif
                    , TITLE, Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Exclamation);

            }
        }

        private static void SelectFragment(Excel.Workbook workbook, string fragment)
        {
            string[] sheet_cell = fragment.Split('!');
            foreach (Excel.Worksheet sheet in workbook.Sheets)
            {
                if (sheet.Name.Equals(sheet_cell[0]) && sheet.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                {
                    sheet.Select();
                    if (sheet_cell.Length <= 1) return;
                    try
                    {
                        sheet.get_Range(sheet_cell[1]).Select();
                    }
                    catch (COMException)
                    {
                        Forms.MessageBox.Show($"Range or Name \"{sheet_cell[1]}\" not found" /*ex.Message + ex.StackTrace*/, TITLE, Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Exclamation);
                    }
                    return;
                }
            }
            Forms.MessageBox.Show($"Sheet \"{sheet_cell[0]}\" not found", TITLE, Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Exclamation);
        }

        private static void SelectFragment(PowerPoint.Presentation ppt, string fragment)
        {
            string[] sheet_cell = fragment.Split('!');

            if (!Int32.TryParse(sheet_cell[0], out int pageNo))
            {
                Forms.MessageBox.Show($"Slide number {sheet_cell[0]} is not a number."
                    , TITLE, Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Exclamation);
                return;
            }
            if (pageNo <= 0 || ppt.Slides.Count < pageNo)
            {
                Forms.MessageBox.Show($"Slide number {pageNo} is out of range. the presentation has only {ppt.Slides.Count} slide[s]."
                    , TITLE, Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Exclamation);
                return;
            }

            try
            {
                ppt.Application.ActiveWindow.View.GotoSlide(pageNo);
                if (sheet_cell.Length <= 1) return;

                string[] names = sheet_cell[1].Split(',');
                PowerPoint.SlideRange slide = ppt.Slides.Range(pageNo);
                foreach (string name in names)
                {
                    new Action(() =>
                    {
                        for (int i = 0; i < slide.Shapes.Count; i++)
                        {
                            PowerPoint.Shape shape = slide.Shapes[1 + i];
                            if (shape.Name.Equals(name))
                            {
                                shape.Select(MsoTriState.msoFalse);
                                return;
                            }
                        }
                        Forms.MessageBox.Show($"Object \"{name}\" not found", TITLE, Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Exclamation);
                    })();
                }
            }
            catch (COMException ex)
            {
                Forms.MessageBox.Show(ex.Message + ex.StackTrace, TITLE, Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Exclamation);
            }
        }

        private static bool Exists(Excel.Names names, string fragment)
        {
            foreach (Excel.Name name in names)
            {
                if (name.NameLocal.Equals(fragment)) return true;
            }
            return false;
        }
    }
}