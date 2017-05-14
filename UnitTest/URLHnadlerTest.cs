using System;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Net;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using UrlHandler;

namespace UnitTest
{
    [TestClass]
    public class URLHnadlerTest
    {
        static string cwd = Directory.GetCurrentDirectory();
        static string PREFIX = UrlHandler.Program.PREFIX;

        [TestMethod]
        public void Test_CWD_Sheet_Range()
        {
            String[] args = { $"{PREFIX}{cwd}\\test Book.xlsx#Sheet2!G3:I4" };
            Program.Main(args);
            _Application excel = Program.GetExcel();
            AssertAddress("Sheet2", "$G$3:$I$4");
            excel.ActiveWorkbook.Close();
        }

        [TestMethod]
        public void Test_CWD_Sheet()
        {
            String[] args = { $"{PREFIX}{cwd}\\test Book.xlsx#Sheet2" };
            Program.Main(args);
            _Application excel = Program.GetExcel();
            AssertAddress("Sheet2", "$C$3:$E$4");
            excel.ActiveWorkbook.Close();
        }

        [TestMethod]
        public void Test_CWD_Sheet_OneCell()
        {
            String[] args = { $"{PREFIX}{cwd}\\test Book.xlsx#Sheet2!G3" };
            Program.Main(args);
            _Application excel = Program.GetExcel();
            AssertAddress("Sheet2", "$G$3");
            excel.ActiveWorkbook.Close();
        }

        [TestMethod]
        [Ignore]
        public void Test_CWD_Sheets_OneCell()
        {
            String[] args = { $"{PREFIX}{cwd}\\test Book.xlsx#Sheet1:Sheet2!D1" };
            Program.Main(args);
            _Application excel = Program.GetExcel();
            AssertAddress("Sheet1:Sheet2", "$D$1");
            excel.ActiveWorkbook.Close();
        }

        [TestMethod]
        public void Test_CWD_Sheet_TwoCells()
        {
            String[] args = { $"{PREFIX}{cwd}\\test Book.xlsx#Sheet2!G4,I3" };
            Program.Main(args);
            _Application excel = Program.GetExcel();
            AssertAddress("Sheet2", "$G$4,$I$3");
            excel.ActiveWorkbook.Close();
        }

        [TestMethod]
        public void Test_CWD_Name()
        {
            String[] args = { $"{PREFIX}{cwd}\\test Book.xlsx#TARGET" };
            Program.Main(args);
            _Application excel = Program.GetExcel();
            AssertAddress("Sheet2", "$G$6:$I$7");
            excel.ActiveWorkbook.Close();
        }

        [TestMethod]
        public void Test_CWD_Name_URLEncoded()
        {
            String[] args = { $"{PREFIX}{cwd}\\test%20Book.xlsx#TARGET" };
            Program.Main(args);
            _Application excel = Program.GetExcel();
            AssertAddress("Sheet2", "$G$6:$I$7");
            excel.ActiveWorkbook.Close();
        }

        [TestMethod]
        public void Test_CWD_WrongName()
        {
            StartCaptureMessage(UrlHandler.Program.TITLE);
            String[] args = { $"{PREFIX}{cwd}\\invalid Book.xlsx#TARGET" };
            Program.Main(args);
            _Application excel = Program.GetExcel();
            Assert.IsTrue(StopCaptureMessage().StartsWith(
                $"申し訳ございません。{cwd}\\invalid Book.xlsxが見つかりません。名前が変更されたか、移動や削除が行われた可能性があります。"));
            // excel.ActiveWorkbook.Close();
        }

        [TestMethod]
        public void Test_CIFS_Name()
        {
            StartCIFS();
            String[] args = { $"{PREFIX}\\\\localhost\\testShare\\test Book.xlsx#TARGET" };
            Program.Main(args);
            _Application excel = Program.GetExcel();
            AssertAddress("Sheet2", "$G$6:$I$7");
            excel.ActiveWorkbook.Close();
            StopCIFS();
        }
        
        [TestMethod]
        public void Test_CIFS_Name_URLEncoded()
        {
            StartCIFS();
            String[] args = { $"{PREFIX}\\\\localhost\\testShare\\test%20Book.xlsx#TARGET" };
            Program.Main(args);
            _Application excel = Program.GetExcel();
            AssertAddress("Sheet2", "$G$6:$I$7");
            excel.ActiveWorkbook.Close();
            StopCIFS();
        }

        [TestMethod]
        public void Test_CIFS_AdminShare()
        {
            string basename = cwd.Substring(3);
            String[] args = { $"{PREFIX}\\\\localhost\\c$\\{basename}\\test Book.xlsx#TARGET" };
            Program.Main(args);
            _Application excel = Program.GetExcel();
            AssertAddress("Sheet2", "$G$6:$I$7");
            excel.ActiveWorkbook.Close();
        }

        [TestMethod]
        public void Test_CIFS_AdminShare_URLEncoded()
        {
            string basename = cwd.Substring(3);
            String[] args = { $"{PREFIX}\\\\localhost\\c$\\{basename}\\test%20Book.xlsx#TARGET" };
            Program.Main(args);
            _Application excel = Program.GetExcel();
            AssertAddress("Sheet2", "$G$6:$I$7");
            excel.ActiveWorkbook.Close();
        }

        [TestMethod]
        public void Test_WebURLEncoded_Name()
        {
            try
            {
                StartWebServer();
                String[] args = { $"{PREFIX}http://localhost:34567/test%20Book.xlsx#TARGET" };
                Program.Main(args);
                AssertAddress("Sheet2", "$G$6:$I$7");
            }
            finally
            {
                StopWebServer();
                _Application excel = Program.GetExcel();
                excel.ActiveWorkbook.Close();
            }
        }

        [TestMethod]
        public void Test_WebwithSpaces_Name()
        {
            try
            {
                StartWebServer();
                String[] args = { $"{PREFIX}http://localhost:34567/test", "Book.xlsx#TARGET" };
                Program.Main(args);
                AssertAddress("Sheet2", "$G$6:$I$7");
            }
            finally
            {
                StopWebServer();
                _Application excel = Program.GetExcel();
                excel.ActiveWorkbook.Close();
            }
        }

        [TestMethod]
        public void Test_Web_SheetNameWithSpaces_Range()
        {
            try
            {
                StartWebServer();
                String[] args = { $"{PREFIX}http://localhost:34567/test%20Book.xlsx#Sheet", "with", "spaces!A1:B2" };
                Program.Main(args);
                AssertAddress("Sheet with spaces", "$A$1:$B$2");
            }
            finally
            {
                StopWebServer();
                _Application excel = Program.GetExcel();
                excel.ActiveWorkbook.Close();
            }
        }

        [TestMethod]
        public void Test_CWD_SheetNameWithSpaces_Range()
        {
            String[] args = { $"{PREFIX}{cwd}\\test Book.xlsx#Sheet with spaces!G3:I4" };
            Program.Main(args);
            _Application excel = Program.GetExcel();
            AssertAddress("Sheet with spaces", "$G$3:$I$4");
            excel.ActiveWorkbook.Close();
        }

        [TestMethod]
        public void Test_CWD_WrongSheet_OneCell()
        {
            StartCaptureMessage(UrlHandler.Program.TITLE);
            String[] args = { $"{PREFIX}{cwd}\\test Book.xlsx#SheetX!G3:I4" };
            Program.Main(args);
            _Application excel = Program.GetExcel();
            Assert.AreEqual("Sheet \"SheetX\" not found", StopCaptureMessage());
            excel.ActiveWorkbook.Close();
        }

        [TestMethod]
        public void Test_CWD_Sheet_WrongCell()
        {
            StartCaptureMessage(UrlHandler.Program.TITLE);
            String[] args = { $"{PREFIX}{cwd}\\test Book.xlsx#Sheet2!XFE1" };
            Program.Main(args);
            _Application excel = Program.GetExcel();
            Assert.AreEqual("Range or Name \"XFE1\" not found", StopCaptureMessage());
            excel.ActiveWorkbook.Close();
        }

        [TestMethod]
        public void Test_CWD_Sheet_WrongName()
        {
            StartCaptureMessage(UrlHandler.Program.TITLE);
            String[] args = { $"{PREFIX}{cwd}\\test Book.xlsx#Sheet2!Noname" };
            Program.Main(args);
            _Application excel = Program.GetExcel();
            Assert.AreEqual("Range or Name \"Noname\" not found", StopCaptureMessage());
            excel.ActiveWorkbook.Close();
        }

        [TestMethod]
        public void Test_CWD_HideenCell()
        {
            String[] args = { $"{PREFIX}{cwd}\\test Book.xlsx#Sheet2!K3:L4" };
            Program.Main(args);
            _Application excel = Program.GetExcel();
            AssertAddress("Sheet2", "$K$3:$L$4");
            excel.ActiveWorkbook.Close();
        }

        [TestMethod]
        public void Test_CWD_HideenSheet()
        {
            StartCaptureMessage(UrlHandler.Program.TITLE);
            String[] args = { $"{PREFIX}{cwd}\\test Book.xlsx#hidden!G3:I4" };
            Program.Main(args);
            _Application excel = Program.GetExcel();
            Assert.AreEqual("Sheet \"hidden\" not found", StopCaptureMessage());
            excel.ActiveWorkbook.Close();
        }

        [TestMethod]
        public void Test_CWD_shrinkedCell()
        {
            String[] args = { $"{PREFIX}{cwd}\\test Book.xlsx#Sheet2!G9:I10" };
            Program.Main(args);
            _Application excel = Program.GetExcel();
            AssertAddress("Sheet2", "$G$9:$I$10");
            excel.ActiveWorkbook.Close();
        }

        [TestMethod]
        public void Test_pptx_CWD_WrongSheet()
        {
            StartCaptureMessage(UrlHandler.Program.TITLE);
            String[] args = { $"{PREFIX}{cwd}\\test Slides.pptx#5" };
            Program.Main(args);
            PowerPoint._Application ppt = Program.GetPowerPoint();
            Assert.AreEqual("Slide number 5 is out of range. the presentation has only 3 slide[s].", StopCaptureMessage());
            ppt.ActivePresentation.Close();
        }

        [TestMethod]
        public void Test_pptx_CWD_PageNoParseError()
        {
            StartCaptureMessage(UrlHandler.Program.TITLE);
            String[] args = { $"{PREFIX}{cwd}\\test Slides.pptx#Noname" };
            Program.Main(args);
            PowerPoint._Application ppt = Program.GetPowerPoint();
            Assert.AreEqual("Slide number Noname is not a number.", StopCaptureMessage());
            ppt.ActivePresentation.Close();
        }

        [TestMethod]
        public void Test_pptx_CWD_Sheet()
        {
            String[] args = { $"{PREFIX}{cwd}\\test Slides.pptx#2" };
            Program.Main(args);
            PowerPoint._Application ppt = Program.GetPowerPoint();
            Assert.AreEqual(2, ppt.ActiveWindow.View.Slide.slideIndex);
            ppt.ActivePresentation.Close();
        }

        [TestMethod]
        public void Test_pptx_CWD_Sheet_Objects()
        {
            String[] args = { $"{PREFIX}{cwd}\\test Slides.pptx#2!Rounded Rectangle 4,Rounded Rectangle 5" };
            Program.Main(args);
            PowerPoint._Application ppt = Program.GetPowerPoint();
            Assert.AreEqual(2, ppt.ActiveWindow.View.Slide.slideIndex);
            PowerPoint.Selection actual = ppt.ActiveWindow.Selection;
            Assert.AreEqual("Rounded Rectangle 4", actual.ShapeRange[1].Name);
            Assert.AreEqual("Rounded Rectangle 5", actual.ShapeRange[2].Name);
            ppt.ActivePresentation.Close();
        }

        [TestMethod]
        public void Test_pptx_CWD_Sheet_WrongObjects()
        {
            StartCaptureMessage(UrlHandler.Program.TITLE);
            String[] args = { $"{PREFIX}{cwd}\\test Slides.pptx#2!Rectangle 3" };
            Program.Main(args);
            PowerPoint._Application ppt = Program.GetPowerPoint();
            Assert.AreEqual(2, ppt.ActiveWindow.View.Slide.slideIndex);
            PowerPoint.Selection actual = ppt.ActiveWindow.Selection;
            Assert.AreEqual("Object \"Rectangle 3\" not found", StopCaptureMessage());
            ppt.ActivePresentation.Close();
        }

        [TestMethod]
        public void Test_FILE_Sheet()
        {
            String[] args = { $"{PREFIX}file:///{cwd}\\test%20Book.xlsx#Sheet2!G3" };
            Program.Main(args);
            _Application excel = Program.GetExcel();
            AssertAddress("Sheet2", "$G$3");
            excel.ActiveWorkbook.Close();
        }

        private static void AssertAddress(string expectedSheet, string expectedCell)
        {
            _Application excel = Program.GetExcel();
            Assert.IsTrue(excel.Selection is Range);
            Assert.AreEqual(expectedSheet, ((Range)(excel.Selection)).Worksheet.Name);
            Assert.AreEqual(expectedCell, ((Range)(excel.Selection)).AddressLocal);
        }

        private static void StartCIFS()
        {
            ProcessStartInfo psi = new ProcessStartInfo()
            {
                Verb = "runas",
                FileName = "net",
                Arguments = $"share testShare=\"{cwd}\""
            };
            Process cmd = Process.Start(psi);
            cmd.WaitForExit();
        }

        private static void StopCIFS()
        {
            ProcessStartInfo psi = new ProcessStartInfo()
            {
                Verb = "runas",
                FileName = "net",
                Arguments = $"share testShare /delete"
            };
            Process cmd = Process.Start(psi);
            cmd.WaitForExit();
        }

        Boolean replied = false;
        HttpListener listener = new HttpListener();

        private void StartWebServer()
        {
            listener.Prefixes.Add("http://localhost:34567/");
            listener.Start();
            Task.Run(() => // only once
            {
                while (!replied)
                {
                    HttpListenerContext context = listener.GetContext();
                    HttpListenerRequest req = context.Request;
                    HttpListenerResponse res = context.Response;
                    string path = cwd + req.Url.LocalPath.Replace("/", "\\");
                    switch (req.HttpMethod)
                    {
                        case "HEAD":
                            res.ContentLength64 = new System.IO.FileInfo(path).Length;
                            res.StatusCode = HttpStatusCode.OK.GetHashCode();
                            break;
                        case "GET":
                            res.StatusCode = HttpStatusCode.OK.GetHashCode();
                            res.ContentLength64 = new System.IO.FileInfo(path).Length;
                            using (Stream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                            {
                                int count = -1; byte[] buffer = new byte[16384];
                                for (; (count = stream.Read(buffer, 0, buffer.Length)) != 0;)
                                    res.OutputStream.Write(buffer, 0, count);
                            }
                            break;
                        case "OPTIONS":
                            res.StatusCode = HttpStatusCode.MethodNotAllowed.GetHashCode();
                            break;
                        default:
                            Console.WriteLine(req.Url.LocalPath);
                            Console.WriteLine(req.HttpMethod);
                            break;
                    }
                    res.Close();
                }
                listener.Stop();
            });
        }

        private void StopWebServer()
        {
            replied = true;
            listener.Abort();
        }

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        [DllImport("user32.dll")]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpStr, int nMaxCount);

        [DllImport("user32.dll")]
        public static extern int GetWindowTextLength(IntPtr hWnd);

        public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumChildWindows(IntPtr hwndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        string actualMessage = null;

        internal void StartCaptureMessage(string expectedTitle)
        {
            Task.Run(() => // only once
            {
                while (!replied)
                {
                    Task.Delay(500);

                    IntPtr hWnd = GetForegroundWindow();

                    if (!IsWindowVisible(hWnd)) continue;

                    int winLen = GetWindowTextLength(hWnd);

                    StringBuilder sb = new StringBuilder(winLen + 1);
                    GetWindowText(hWnd, sb, sb.Capacity);
                    if (expectedTitle != sb.ToString()) continue;

                    GetWindowThreadProcessId(hWnd, out int id);

                    EnumChildWindows(hWnd, (IntPtr _hWnd, IntPtr lParam) =>
                    {
                        StringBuilder _sb = new StringBuilder(GetWindowTextLength(_hWnd) + 1);
                        GetWindowText(_hWnd, _sb, _sb.Capacity);
                        actualMessage = _sb.ToString();
                        return true;
                    }, hWnd);
                    replied = true;
                }
            });
        }

        internal string StopCaptureMessage()
        {
            return actualMessage;
        }
    }
}
