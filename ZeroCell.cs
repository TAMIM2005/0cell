using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ZeroCell
{
    [ComImport, Guid("886D8EEB-8CF2-4446-8D02-C335F478E719"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IPropertyStore
    {
        int GetCount(out uint cProps);
        int GetAt(uint iProp, out PROPERTYKEY pkey);
        int GetValue(ref PROPERTYKEY key, out PROPVARIANT pv);
        int SetValue(ref PROPERTYKEY key, ref PROPVARIANT pv);
        int Commit();
    }

    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    struct PROPERTYKEY
    {
        public Guid fmtid;
        public uint pid;

        public PROPERTYKEY(Guid g, uint p)
        {
            fmtid = g;
            pid = p;
        }
    }

    [StructLayout(LayoutKind.Explicit)]
    struct PROPVARIANT
    {
        [FieldOffset(0)] public ushort vt;
        [FieldOffset(2)] public ushort wReserved1;
        [FieldOffset(4)] public ushort wReserved2;
        [FieldOffset(6)] public ushort wReserved3;
        [FieldOffset(8)] public IntPtr pwszVal;
    }

    static class Program
    {
        [DllImport("propsys.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern int InitPropVariantFromString([MarshalAs(UnmanagedType.LPWStr)] string psz, out PROPVARIANT ppropvar);

        [DllImport("propsys.dll", SetLastError = true)]
        private static extern int PropVariantClear(ref PROPVARIANT pvar);

        [DllImport("shell32.dll", SetLastError = true)]
        private static extern void SetCurrentProcessExplicitAppUserModelID([MarshalAs(UnmanagedType.LPWStr)] string AppID);

        [DllImport("shell32.dll", SetLastError = true)]
        private static extern int SHGetPropertyStoreForWindow(IntPtr handle, ref Guid riid, out IPropertyStore propertyStore);

        [DllImport("shell32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern void SHChangeNotify(int wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr LoadImage(IntPtr hinst, string lpszName, uint uType, int cxDesired, int cyDesired, uint fuLoad);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        private const uint WM_SETICON = 0x0080;
        private static readonly IntPtr ICON_SMALL = new IntPtr(0);
        private static readonly IntPtr ICON_BIG = new IntPtr(1);
        private const uint IMAGE_ICON = 1;
        private const uint LR_LOADFROMFILE = 0x00000010;

        private static readonly Guid IID_IPropertyStore = new Guid("886D8EEB-8CF2-4446-8D02-C335F478E719");
        private static readonly PROPERTYKEY PKEY_AppUserModel_ID = new PROPERTYKEY(new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3"), 5);
        private static readonly PROPERTYKEY PKEY_AppUserModel_RelaunchIconResource = new PROPERTYKEY(new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3"), 2);
        private static readonly PROPERTYKEY PKEY_AppUserModel_RelaunchCommand = new PROPERTYKEY(new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3"), 2);
        private static readonly PROPERTYKEY PKEY_AppUserModel_RelaunchDisplayNameResource = new PROPERTYKEY(new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3"), 4);

        private const string APP_ID = "ZeroCell.ModernSpreadsheet.App";

        private static HttpListener listener;
        private static string appDir;

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                SetCurrentProcessExplicitAppUserModelID(APP_ID);
            }
            catch { }

            appDir = AppDomain.CurrentDomain.BaseDirectory;
            string htmlPath = Path.Combine(appDir, "index.html");
            string icoPath = Path.Combine(appDir, "0C.ico");
            string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;

            if (!File.Exists(htmlPath))
            {
                MessageBox.Show("Could not locate index.html in the application directory.", "0cell Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int port = 49281;
            string baseUrl = string.Format("http://127.0.0.1:{0}/", port);
            StartLocalServer(port);

            string edgePath = @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe";
            if (!File.Exists(edgePath))
            {
                edgePath = @"C:\Program Files\Microsoft\Edge\Application\msedge.exe";
            }

            string targetUrl = baseUrl + "index.html";
            if (args.Length > 0 && File.Exists(args[0]))
            {
                targetUrl += "?file=" + Uri.EscapeDataString(args[0]);
            }

            string userDataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "0cell", "isolated_profile");

            if (File.Exists(edgePath))
            {
                string iconResource = File.Exists(icoPath) ? icoPath + ",0" : (File.Exists(exePath) ? exePath + ",0" : "");

                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = edgePath,
                    Arguments = string.Format(
                        "--app=\"{0}\" --user-data-dir=\"{1}\" --window-size=1360,820 --disable-features=Translate,OptimizationHints --disable-extensions --no-default-browser-check --no-first-run",
                        targetUrl, userDataDir),
                    UseShellExecute = false
                };
                Process proc = Process.Start(psi);

                // Run Window Icon & AppID monitor in background
                Thread iconThread = new Thread(() =>
                {
                    IntPtr hIconBig = IntPtr.Zero;
                    IntPtr hIconSmall = IntPtr.Zero;

                    if (!File.Exists(icoPath) && File.Exists(exePath))
                    {
                        icoPath = exePath;
                    }

                    if (File.Exists(icoPath))
                    {
                        hIconBig = LoadImage(IntPtr.Zero, icoPath, IMAGE_ICON, 48, 48, LR_LOADFROMFILE);
                        hIconSmall = LoadImage(IntPtr.Zero, icoPath, IMAGE_ICON, 16, 16, LR_LOADFROMFILE);
                    }

                    while (proc != null && !proc.HasExited)
                    {
                        Thread.Sleep(200);
                        EnumWindows((hWnd, lParam) =>
                        {
                            if (!IsWindowVisible(hWnd)) return true;

                            StringBuilder sbClass = new StringBuilder(256);
                            GetClassName(hWnd, sbClass, 256);
                            string className = sbClass.ToString();

                            if (className == "Chrome_WidgetWin_1")
                            {
                                StringBuilder sbTitle = new StringBuilder(256);
                                GetWindowText(hWnd, sbTitle, 256);
                                string title = sbTitle.ToString();

                                if (string.IsNullOrEmpty(title) ||
                                    title.IndexOf("0cell", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    title.IndexOf("Spreadsheet", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    title.IndexOf("127.0.0.1", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    try
                                    {
                                        Guid guid = IID_IPropertyStore;
                                        IPropertyStore store;
                                        if (SHGetPropertyStoreForWindow(hWnd, ref guid, out store) == 0 && store != null)
                                        {
                                            PROPERTYKEY keyId = PKEY_AppUserModel_ID;
                                            PROPVARIANT pvId;
                                            if (InitPropVariantFromString(APP_ID, out pvId) == 0)
                                            {
                                                store.SetValue(ref keyId, ref pvId);
                                                PropVariantClear(ref pvId);
                                            }

                                            if (!string.IsNullOrEmpty(iconResource))
                                            {
                                                PROPERTYKEY keyIcon = PKEY_AppUserModel_RelaunchIconResource;
                                                PROPVARIANT pvIcon;
                                                if (InitPropVariantFromString(iconResource, out pvIcon) == 0)
                                                {
                                                    store.SetValue(ref keyIcon, ref pvIcon);
                                                    PropVariantClear(ref pvIcon);
                                                }
                                            }

                                            PROPERTYKEY keyCmd = PKEY_AppUserModel_RelaunchCommand;
                                            PROPVARIANT pvCmd;
                                            if (InitPropVariantFromString("\"" + exePath + "\"", out pvCmd) == 0)
                                            {
                                                store.SetValue(ref keyCmd, ref pvCmd);
                                                PropVariantClear(ref pvCmd);
                                            }

                                            PROPERTYKEY keyName = PKEY_AppUserModel_RelaunchDisplayNameResource;
                                            PROPVARIANT pvName;
                                            if (InitPropVariantFromString("0cell", out pvName) == 0)
                                            {
                                                store.SetValue(ref keyName, ref pvName);
                                                PropVariantClear(ref pvName);
                                            }

                                            store.Commit();
                                        }
                                    }
                                    catch { }

                                    if (hIconBig != IntPtr.Zero)
                                        SendMessage(hWnd, WM_SETICON, ICON_BIG, hIconBig);
                                    if (hIconSmall != IntPtr.Zero)
                                        SendMessage(hWnd, WM_SETICON, ICON_SMALL, hIconSmall);
                                }
                            }
                            return true;
                        }, IntPtr.Zero);
                    }
                });
                iconThread.IsBackground = true;
                iconThread.Start();

                if (proc != null)
                {
                    proc.WaitForExit();
                }
            }
            else
            {
                Process.Start(new ProcessStartInfo(targetUrl) { UseShellExecute = true });
            }
        }

        private static void StartLocalServer(int port)
        {
            try
            {
                listener = new HttpListener();
                listener.Prefixes.Add(string.Format("http://127.0.0.1:{0}/", port));
                listener.Start();

                Thread serverThread = new Thread(() =>
                {
                    while (listener.IsListening)
                    {
                        try
                        {
                            var context = listener.GetContext();
                            ThreadPool.QueueUserWorkItem((ctx) => HandleRequest((HttpListenerContext)ctx), context);
                        }
                        catch { }
                    }
                });
                serverThread.IsBackground = true;
                serverThread.Start();
            }
            catch { }
        }

        private static void HandleRequest(HttpListenerContext context)
        {
            try
            {
                string rawUrl = context.Request.Url.AbsolutePath.TrimStart('/');
                if (string.IsNullOrEmpty(rawUrl)) rawUrl = "index.html";

                string filePath = Path.Combine(appDir, rawUrl.Replace('/', Path.DirectorySeparatorChar));

                if (File.Exists(filePath))
                {
                    byte[] bytes = File.ReadAllBytes(filePath);
                    string ext = Path.GetExtension(filePath).ToLowerInvariant();
                    string mime = "application/octet-stream";

                    if (ext == ".html") mime = "text/html; charset=utf-8";
                    else if (ext == ".js") mime = "application/javascript; charset=utf-8";
                    else if (ext == ".css") mime = "text/css; charset=utf-8";
                    else if (ext == ".ico") mime = "image/x-icon";
                    else if (ext == ".webp") mime = "image/webp";
                    else if (ext == ".png") mime = "image/png";
                    else if (ext == ".json") mime = "application/json; charset=utf-8";
                    else if (ext == ".svg") mime = "image/svg+xml";

                    context.Response.ContentType = mime;
                    context.Response.ContentLength64 = bytes.Length;
                    context.Response.OutputStream.Write(bytes, 0, bytes.Length);
                }
                else
                {
                    context.Response.StatusCode = 404;
                }
            }
            catch { }
            finally
            {
                try { context.Response.Close(); } catch { }
            }
        }
    }
}
