using NativeCode;
using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.Linq;
using System.Reflection;
using System.Reflection.Metadata;
using System.Runtime;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Win32;
using Windows.ApplicationModel;
using Windows.ApplicationModel.DataTransfer;
using Windows.ApplicationModel.DataTransfer.ShareTarget;
using Windows.Storage;
using Windows.Storage.Streams;
using Windows.Storage.Search;
using Windows.Storage.Pickers;
using Windows.Storage.Provider;
using Windows.UI;
using Windows.UI.Core;
using Windows.UI.Text;
using WinRT;
using System.Globalization;

namespace CharmsBarPort
{
    #region Cursor info
    public static class CursorExtensions
    {

        [StructLayout(LayoutKind.Sequential)]
        struct PointStruct
        {
            public Int32 x;
            public Int32 y;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct CursorInfoStruct
        {
            /// <summary> The structure size in bytes that must be set via calling Marshal.SizeOf(typeof(CursorInfoStruct)).</summary>
            public Int32 cbSize;
            /// <summary> The cursor state: 0 == hidden, 1 == showing, 2 == suppressed (is supposed to be when finger touch is used, but in practice finger touch results in 0, not 2)</summary>
            public Int32 flags;
            /// <summary> A handle to the cursor. </summary>
            public IntPtr hCursor;
            /// <summary> The cursor screen coordinates.</summary>
            public PointStruct pt;
        }

        /// <summary> Must initialize cbSize</summary>
        [DllImport("user32.dll")]
        static extern bool GetCursorInfo(ref CursorInfoStruct pci);
        public static bool IsVisible(this System.Windows.Forms.Cursor cursor)
        {
            CursorInfoStruct pci = new CursorInfoStruct();
            pci.cbSize = Marshal.SizeOf(typeof(CursorInfoStruct));
            GetCursorInfo(ref pci);
            // const Int32 hidden = 0x00;
            const Int32 showing = 0x01;
            // const Int32 suppressed = 0x02;
            bool isVisible = ((pci.flags & showing) != 0);
            return isVisible;
        }

    }
    #endregion Cursor info
    public sealed partial class CharmsBar : Window
    {
        //the Share Charm is all fired up for its new iteration!
        [ComImport]
        [Guid("3A3DCD6C-3EAB-43DC-BCDE-45671CE800C8")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IDataTransferManagerInterop
        {
            IntPtr GetForWindow([In] IntPtr appWindow, [In] ref Guid riid);
            void ShowShareUIForWindow(IntPtr appWindow);
        }

        public static DataTransferManager GetDataTransferManager(IntPtr appWindow)
        {
            IDataTransferManagerInterop interop = DataTransferManager.As<IDataTransferManagerInterop>();
            Guid id = new Guid(0xa5caee9b, 0x8708, 0x49d1, 0x8d, 0x36, 0x67, 0xd2, 0x5a, 0x8d, 0xa0, 0x0c);
            IntPtr result;
            result = interop.GetForWindow(appWindow, id);
            DataTransferManager dataTransferManager = MarshalInterface<DataTransferManager>.FromAbi(result);
            return (dataTransferManager);
        }

        DataTransferManagerHelper dtmHelper = null;
        List<IStorageItem> filesToShare = null;

        static readonly Guid _dtm_iid =
            new Guid(0xa5caee9b, 0x8708, 0x49d1, 0x8d, 0x36, 0x67, 0xd2, 0x5a, 0x8d, 0xa0, 0x0c);

        private void EnsureDataTransferManager()
        {
            if (this.dtmHelper == null)
            {
                IntPtr windowHandle = new WindowInteropHelper(this).Handle;
                this.dtmHelper = new DataTransferManagerHelper(windowHandle);
                this.dtmHelper.DataTransferManager.DataRequested += this.OnDataRequested;
            }
        }

        protected void OnDataRequested(DataTransferManager sender, DataRequestedEventArgs e)
        {
            e.Request.Data.Properties.Title = " ";
            e.Request.Data.Properties.Description = " ";
            e.Request.Data.SetText(" ");
        }
        public int rctLeft = 0;
        public int rctTop = 0;
        public bool findDevices = false;
        public bool openSettings = false;
        public int findTimer = 0;
        BrushConverter converter = new();
        Window CharmsClock = new CharmsClock();
        Window CharmsMenu = new CharmsMenu();
        public bool holder = false;
        public int dasBoot = 0;
        static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
        static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        static readonly IntPtr HWND_TOP = new IntPtr(0);
        static readonly IntPtr HWND_BOTTOM = new IntPtr(1);
        const UInt32 SWP_NOSIZE = 0x0001;
        const UInt32 SWP_NOMOVE = 0x0002;
        const UInt32 SWP_SHOWWINDOW = 0x4000;
        const UInt32 TOPMOST_FLAGS = SWP_SHOWWINDOW | SWP_NOMOVE | SWP_NOSIZE;
        public bool preventReload = false;
        public int blockRepeating = 0;
        public int cursorStay = 0;
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hwind, int cmd);

        public bool charmsMenuOpen = false;
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        public static IntPtr GetWindowHandle(Window window)
        {
            return new WindowInteropHelper(window).Handle;
        }
        private void BringToFront(Process pTemp)
        {
            SetForegroundWindow(pTemp.MainWindowHandle);
        }

        public bool forceClose = false;
        public bool charmsFade = false;
        public int activeScreen = 0;
        public bool swipeIn = false;
        public bool keyboardShortcut = false;
        public bool charmsAppear = false;
        public bool charmsUse = false;
        public int charmsTimer = 0;
        public int charmsWait = 0;
        public bool WinCharmUse = false;
        public int charmsDelay = 100;
        public int myCharmsDelay = 100; // Desktop delay: you can customize the delay option through Regedit.
        public int charmsDelay2 = 100; // Metro delay: you can customize the delay option through Regedit.
        public int charmsReturn = 0; //to fix a problem where you can't swipe back in after it's gone
        public int activeIcon = 2;
        public bool mouseIn = false;
        public bool twoInputs = false;
        public int waitTimer = 0;
        public int keyboardTimer = 0;
        public bool charmsActivate = false;
        public double IHOb = 1.0;
        public bool escKey = false;
        public bool pokeCharms = false;
        public bool usingTouch = false;
        public bool isMetro = false; //Metro apps use their unique ways for stuff
        public bool isGui = false; //Fixing a problem where it appears behind the taskbar.
        public bool searchHover = false;
        public bool shareHover = false;
        public bool winHover = false;
        public bool devicesHover = false;
        public bool settingsHover = false;
        public IntPtr mWnd = GetForegroundWindow();

        public bool searchActive = false;
        public bool shareActive = false;
        public bool winActive = false;
        public bool devicesActive = false;
        public bool settingsActive = false;

        //Supports Windows 8.1 / Windows 10 registry hacks!
        public string customDelay = "100";

        //For the animations!
        public bool useAnimations = false;
        public double winStretch = 80.31;
        public int dasSlide = 0;
        public int scrollSearch = 200;
        public int scrollShare = 150;
        public int scrollWin = 100;
        public int scrollDevices = 150;
        public int scrollSettings = 200;

        public int textSearch = 170;
        public int textShare = 150;
        public int textWin = 100;
        public int textDevices = 150;
        public int textSettings = 200;

        //mouse
        public bool ignoreMouseIn = false;
        public bool outofTime = false;
        public int numVal = 0;
        public int numVal2 = 0;

        //multi-monitor
        public bool dasSwiper = false;

        public int mainwidth = 0;
        public int mainheight = 0;
        public int mainX = 0;
        public int twowidth = 0;
        public int twoheight = 0;
        public int twoX = 0;
        public int threewidth = 0;
        public int threeheight = 0;
        public int threeX = 0;
        public int fourwidth = 0;
        public int fourheight = 0;
        public int fourX = 0;
        public int fivewidth = 0;
        public int fiveheight = 0;
        public int fiveX = 0;
        public int sixwidth = 0;
        public int sixheight = 0;
        public int sixX = 0;
        public int sevenwidth = 0;
        public int sevenheight = 0;
        public int sevenX = 0;
        public int eightwidth = 0;
        public int eightheight = 0;
        public int eightX = 0;
        public int ninewidth = 0;
        public int nineheight = 0;
        public int nineX = 0;
        public int tenwidth = 0;
        public int tenheight = 0;
        public int tenX = 0;
        public int elevenwidth = 0;
        public int elevenheight = 0;
        public int elevenX = 0;

        public int screenwidth = 0; //just to make things more reliable
        public int screenheight = 0; //just to make things more reliable
        public int screenX = 0; //just to make things more reliable

        public CharmsBar()
        {
            Topmost = true;
            ShowInTaskbar = false;
            WindowStyle = WindowStyle.None;
            ResizeMode = ResizeMode.NoResize;
            AllowsTransparency = true;
            Height = SystemParameters.PrimaryScreenHeight;
            Width = 86;
            WindowStartupLocation = WindowStartupLocation.Manual;
            Left = Screen.PrimaryScreen.Bounds.Width - 86;
            Top = 0;
            var brush = (Brush)converter.ConvertFromString("#00111111");
            Background = brush;
            Opacity = 0.000;
            SystemParameters.StaticPropertyChanged += this.SystemParameters_StaticPropertyChanged;
            this.Loaded += ControlLoaded;
            CharmsClock.Hide();
            CharmsMenu.Hide();
            this.Closing += new System.ComponentModel.CancelEventHandler(MainWindow_Closing);
            this.KeyDown += new System.Windows.Input.KeyEventHandler(MainWindow_KeyDown);
            System.Windows.Forms.Application.ThreadException += new ThreadExceptionEventHandler(CharmsBar.Form1_UIThreadException);
            InitializeComponent();
        }

        void MainWindow_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (charmsUse == true)
            {
                if (this.IsActive == true && keyboardShortcut == false)
                {
                    keyboardShortcut = true;
                    escKey = true;
                }
            }
        }
        void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
        }

private string GetActiveWindowTitle()
{
    const int nChars = 256;
    StringBuilder Buff = new StringBuilder(nChars);
    IntPtr handle = GetForegroundWindow();

    if (GetWindowText(handle, Buff, nChars) > 0)
    {
        return Buff.ToString();
    }
    return null;
}

   [DllImport("user32.dll",CharSet=CharSet.Auto, CallingConvention=CallingConvention.StdCall)]
   public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);
   //Mouse actions
   private const int MOUSEEVENTF_LEFTDOWN = 0x02;
   private const int MOUSEEVENTF_LEFTUP = 0x04;
   private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
   private const int MOUSEEVENTF_RIGHTUP = 0x10;

        [DllImport("User32.dll")]
        private static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)] 
    static extern IntPtr FindWindow(string lpClassName, string lpWindowName); 
    [DllImport("user32.dll", EntryPoint = "SendMessage", SetLastError = true)] 
    static extern IntPtr SendMessage(IntPtr hWnd, Int32 Msg, IntPtr wParam, IntPtr lParam); 

    const int WM_COMMAND = 0x111; 
    const int MIN_ALL = 419; 
    const int MIN_ALL_UNDO = 416; 

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern int GetSystemMetrics(int nIndex);

        public static bool IsTouchEnabled()
        {
            const int MAXTOUCHES_INDEX = 95;
            int maxTouches = GetSystemMetrics(MAXTOUCHES_INDEX);

            return maxTouches > 0;
        }

        protected override void OnActivated(EventArgs e)
        {
            if (this.IsActive == true && charmsUse == true)
            {
                if (useAnimations == false)
                {
                    this.Opacity = 1.0;
                    CharmsClock.Opacity = 1.0;
                    if (noClocks.Content == "-1" || noClocks.Content == "0")
                    {
                        CharmsClock.Show();
                    }
                }
            }
            base.OnActivated(e);
        }

        protected override void OnDeactivated(EventArgs e)
        {
            if (useAnimations == false)
            {
                swipeIn = false;
                twoInputs = false;
                keyboardShortcut = false;
                charmsAppear = false;
                charmsActivate = false;
                pokeCharms = false;
                if (useAnimations == false)
                {
                    this.Opacity = 0.000;
                    CharmsClock.Opacity = 0.000;

                    var brush = (Brush)converter.ConvertFromString("#00111111");
                    Background = brush;
                }
                mouseIn = false;

                SearchDown.Visibility = Visibility.Hidden;
                ShareDown.Visibility = Visibility.Hidden;
                WinDown.Visibility = Visibility.Hidden;
                DevicesDown.Visibility = Visibility.Hidden;
                SettingsDown.Visibility = Visibility.Hidden;

                SearchText.Visibility = Visibility.Hidden;
                ShareText.Visibility = Visibility.Hidden;
                WinText.Visibility = Visibility.Hidden;
                DevicesText.Visibility = Visibility.Hidden;
                SettingsText.Visibility = Visibility.Hidden;

                SearchCharm.Visibility = Visibility.Hidden;
                ShareCharm.Visibility = Visibility.Hidden;
                MetroColor.Visibility = Visibility.Hidden;
                DevicesCharm.Visibility = Visibility.Hidden;
                SettingsCharm.Visibility = Visibility.Hidden;

                SearchCharmInactive.Visibility = Visibility.Visible;
                ShareCharmInactive.Visibility = Visibility.Visible;
                NoColor.Visibility = Visibility.Visible;
                if (vn4.Content != "0" && vn4.Content != "-1")
                {
                    DevicesCharmInactive.Visibility = Visibility.Hidden;
                }
                else
                {
                    DevicesCharmInactive.Visibility = Visibility.Visible;
                }

                if (vn3.Content != "0" && vn3.Content != "-1")
                {
                    SettingsCharmInactive.Visibility = Visibility.Hidden;
                }
                else
                {
                    SettingsCharmInactive.Visibility = Visibility.Visible;
                }

                SearchHover.Visibility = Visibility.Hidden;
                ShareHover.Visibility = Visibility.Hidden;
                WinHover.Visibility = Visibility.Hidden;
                DevicesHover.Visibility = Visibility.Hidden;
                SettingsHover.Visibility = Visibility.Hidden;

                charmsUse = false;
            }

            if (charmsTimer == 0 && keyboardShortcut == false)
            {
                swipeIn = false;
                twoInputs = false;
                keyboardShortcut = false;
                charmsAppear = false;
                charmsActivate = false;
                pokeCharms = false;
                if (useAnimations == false)
                {
                    this.Opacity = 0.000;
                    CharmsClock.Opacity = 0.000;

                    var brush = (Brush)converter.ConvertFromString("#00111111");
                    Background = brush;
                }
                mouseIn = false;

                SearchDown.Visibility = Visibility.Hidden;
                ShareDown.Visibility = Visibility.Hidden;
                WinDown.Visibility = Visibility.Hidden;
                DevicesDown.Visibility = Visibility.Hidden;
                SettingsDown.Visibility = Visibility.Hidden;

                SearchText.Visibility = Visibility.Hidden;
                ShareText.Visibility = Visibility.Hidden;
                WinText.Visibility = Visibility.Hidden;
                DevicesText.Visibility = Visibility.Hidden;
                SettingsText.Visibility = Visibility.Hidden;

                SearchCharm.Visibility = Visibility.Hidden;
                ShareCharm.Visibility = Visibility.Hidden;
                MetroColor.Visibility = Visibility.Hidden;
                DevicesCharm.Visibility = Visibility.Hidden;
                SettingsCharm.Visibility = Visibility.Hidden;

                SearchCharmInactive.Visibility = Visibility.Visible;
                ShareCharmInactive.Visibility = Visibility.Visible;
                NoColor.Visibility = Visibility.Visible;
                DevicesCharmInactive.Visibility = Visibility.Visible;
                SettingsCharmInactive.Visibility = Visibility.Visible;

                SearchHover.Visibility = Visibility.Hidden;
                ShareHover.Visibility = Visibility.Hidden;
                WinHover.Visibility = Visibility.Hidden;
                DevicesHover.Visibility = Visibility.Hidden;
                SettingsHover.Visibility = Visibility.Hidden;

                charmsUse = false;
            }

            if (useAnimations == false)
            {
                if (CharmsClock.Opacity == 1.0)
                {
                    CharmsClock.Show();
                }
                else
                {
                    {
                        CharmsClock.Hide();
                    }
                }
            }

            base.OnDeactivated(e);
        }

        protected override void OnClosed(EventArgs e)
        {
            SystemParameters.StaticPropertyChanged -= this.SystemParameters_StaticPropertyChanged;
            base.OnClosed(e);
        }

        public void ControlLoaded(object sender, EventArgs e)
        {
            var wih = new System.Windows.Interop.WindowInteropHelper(this);
            SetWindowPos(wih.Handle, HWND_TOPMOST, 100, 100, 300, 300, TOPMOST_FLAGS);
            _initTimer();
        }

        private void SetBackgroundColor()
        {
            MetroColor.Background = SystemParameters.WindowGlassBrush;
        }

        private void SystemParameters_StaticPropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "WindowGlassBrush")
            {
                this.SetBackgroundColor();
            }
        }

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SystemParametersInfo(uint uiAction, uint uiParam, out bool pvParam, uint fWinIni);

        private static uint SPI_GETCLIENTAREAANIMATION = 0x1042;

        [DllImport("User32")]

        private static extern int keybd_event(byte bVk, byte bScan, uint dwFlags, long dwExtraInfo);

        //for Metro Apps
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowRect(IntPtr hWnd, ref RECT lpRect);
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        private static bool isMetroApp(IntPtr hWnd)
        {
            int nRet;
            // Pre-allocate 256 characters, since this is the maximum class name length.
            StringBuilder ClassName = new StringBuilder(256);
            //Get the window class name
            nRet = GetClassName(hWnd, ClassName, ClassName.Capacity);
            if (nRet != 0)
            {
                return (string.Compare(ClassName.ToString(), "ApplicationFrameWindow", true, CultureInfo.InvariantCulture) == 0);
            }
            else
            {
                return false;
            }
        }

        private static bool isWinGui(IntPtr hWnd)
        {
            int nRet;
            // Pre-allocate 256 characters, since this is the maximum class name length.
            StringBuilder ClassName = new StringBuilder(256);
            //Get the window class name
            nRet = GetClassName(hWnd, ClassName, ClassName.Capacity);
            if (nRet != 0)
            {
                return (string.Compare(ClassName.ToString(), "Windows.UI.Core.CoreWindow", true, CultureInfo.InvariantCulture) == 0);
            }
            else
            {
                return false;
            }
        }

        private async void Search_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (isMetro == true)
            {
                RECT rct = new RECT();
                GetWindowRect(mWnd, ref rct);
                SetCursorPos(rctLeft + 21, rctTop + 14);

                //perform click            
                mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                await Task.Delay(333);
                mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                await Task.Delay(333);
                byte sKey = (byte)KeyInterop.VirtualKeyFromKey(Key.S);
                const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
                const uint KEYEVENTF_KEYUP = 0x0002;
                _ = keybd_event(sKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
                await Task.Delay(100);
                _ = keybd_event(sKey, 0, KEYEVENTF_KEYUP, 0);
            }
            else
            {
                searchActive = false;
                byte winKey = (byte)KeyInterop.VirtualKeyFromKey(Key.LWin);
                byte sKey = (byte)KeyInterop.VirtualKeyFromKey(Key.S);
                const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
                const uint KEYEVENTF_KEYUP = 0x0002;
                if (this.IsActive == true)
                {
                    _ = keybd_event(winKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
                    _ = keybd_event(sKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
                    _ = keybd_event(winKey, 0, KEYEVENTF_KEYUP, 0);
                    _ = keybd_event(sKey, 0, KEYEVENTF_KEYUP, 0);
                }

                swipeIn = false;
                keyboardShortcut = false;
                charmsAppear = false;
                charmsUse = false;
                charmsActivate = false;
                pokeCharms = false;

                if (useAnimations == false)
                {
                    this.Opacity = 0.000;
                    CharmsClock.Opacity = 0.000;

                    var brush = (Brush)converter.ConvertFromString("#00111111");
                    Background = brush;
                }
                mouseIn = false;

                SearchDown.Visibility = Visibility.Hidden;
                ShareDown.Visibility = Visibility.Hidden;
                WinDown.Visibility = Visibility.Hidden;
                DevicesDown.Visibility = Visibility.Hidden;
                SettingsDown.Visibility = Visibility.Hidden;

                SearchText.Visibility = Visibility.Hidden;
                ShareText.Visibility = Visibility.Hidden;
                WinText.Visibility = Visibility.Hidden;
                DevicesText.Visibility = Visibility.Hidden;
                SettingsText.Visibility = Visibility.Hidden;

                SearchCharm.Visibility = Visibility.Hidden;
                ShareCharm.Visibility = Visibility.Hidden;
                MetroColor.Visibility = Visibility.Hidden;
                DevicesCharm.Visibility = Visibility.Hidden;
                SettingsCharm.Visibility = Visibility.Hidden;

                SearchCharmInactive.Visibility = Visibility.Visible;
                ShareCharmInactive.Visibility = Visibility.Visible;
                NoColor.Visibility = Visibility.Visible;
                DevicesCharmInactive.Visibility = Visibility.Visible;
                SettingsCharmInactive.Visibility = Visibility.Visible;
            }
        }
        private async void Share_MouseUp(object sender, RoutedEventArgs e)
        {
            if (isMetro == true)
            {
                RECT rct = new RECT();
                GetWindowRect(mWnd, ref rct);
                SetCursorPos(rctLeft + 21, rctTop + 14);

                //perform click            
                mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                await Task.Delay(333);
                mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                await Task.Delay(333);
                byte hKey = (byte)KeyInterop.VirtualKeyFromKey(Key.H);
                const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
                const uint KEYEVENTF_KEYUP = 0x0002;
                _ = keybd_event(hKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
                await Task.Delay(100);
                _ = keybd_event(hKey, 0, KEYEVENTF_KEYUP, 0);
            }
            else
            {
                var hWnd = new WindowInteropHelper(this).Handle;
                IDataTransferManagerInterop interop =
                Windows.ApplicationModel.DataTransfer.DataTransferManager.As
                    <IDataTransferManagerInterop>();

                IntPtr result = interop.GetForWindow(hWnd, _dtm_iid);
                var dataTransferManager = WinRT.MarshalInterface
                    <Windows.ApplicationModel.DataTransfer.DataTransferManager>.FromAbi(result);

                dataTransferManager.DataRequested += (sender, args) =>
                {
                    args.Request.Data.Properties.Title = " ";
                    args.Request.Data.SetText("WinRT.Interop.WindowNative.GetWindowHandle(this)");
                    args.Request.Data.RequestedOperation =
                        Windows.ApplicationModel.DataTransfer.DataPackageOperation.Copy;
                };

                if (this.IsActive == true)
                {
                    interop.ShowShareUIForWindow(hWnd);
                }

                swipeIn = false;
                keyboardShortcut = false;
                charmsAppear = false;
                charmsUse = false;
                charmsActivate = false;
                pokeCharms = false;

                if (useAnimations == false)
                {
                    this.Opacity = 0.000;
                    CharmsClock.Opacity = 0.000;

                    var brush = (Brush)converter.ConvertFromString("#00111111");
                    Background = brush;
                }

                mouseIn = false;

                SearchDown.Visibility = Visibility.Hidden;
                ShareDown.Visibility = Visibility.Hidden;
                WinDown.Visibility = Visibility.Hidden;
                DevicesDown.Visibility = Visibility.Hidden;
                SettingsDown.Visibility = Visibility.Hidden;

                SearchText.Visibility = Visibility.Hidden;
                ShareText.Visibility = Visibility.Hidden;
                WinText.Visibility = Visibility.Hidden;
                DevicesText.Visibility = Visibility.Hidden;
                SettingsText.Visibility = Visibility.Hidden;

                SearchCharm.Visibility = Visibility.Hidden;
                ShareCharm.Visibility = Visibility.Hidden;
                MetroColor.Visibility = Visibility.Hidden;
                DevicesCharm.Visibility = Visibility.Hidden;
                SettingsCharm.Visibility = Visibility.Hidden;

                SearchCharmInactive.Visibility = Visibility.Visible;
                ShareCharmInactive.Visibility = Visibility.Visible;
                NoColor.Visibility = Visibility.Visible;
                DevicesCharmInactive.Visibility = Visibility.Visible;
                SettingsCharmInactive.Visibility = Visibility.Visible;
            }
        }

        private void Win_MouseUp(object sender, MouseButtonEventArgs e)
        {
            byte winKey = (byte)KeyInterop.VirtualKeyFromKey(Key.LWin);
            const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
            const uint KEYEVENTF_KEYUP = 0x0002;
            _ = keybd_event(winKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
            _ = keybd_event(winKey, 0, KEYEVENTF_KEYUP, 0);
            swipeIn = false;
            keyboardShortcut = false;
            charmsAppear = false;
            charmsUse = false;
            charmsActivate = false;
            pokeCharms = false;
            if (useAnimations == false)
            {
                this.Opacity = 0.000;
                CharmsClock.Opacity = 0.000;
                var brush = (Brush)converter.ConvertFromString("#00111111");
                Background = brush;
            }

            mouseIn = false;

            SearchDown.Visibility = Visibility.Hidden;
            ShareDown.Visibility = Visibility.Hidden;
            WinDown.Visibility = Visibility.Hidden;
            DevicesDown.Visibility = Visibility.Hidden;
            SettingsDown.Visibility = Visibility.Hidden;

            SearchText.Visibility = Visibility.Hidden;
            ShareText.Visibility = Visibility.Hidden;
            WinText.Visibility = Visibility.Hidden;
            DevicesText.Visibility = Visibility.Hidden;
            SettingsText.Visibility = Visibility.Hidden;

            SearchCharm.Visibility = Visibility.Hidden;
            ShareCharm.Visibility = Visibility.Hidden;
            MetroColor.Visibility = Visibility.Hidden;
            DevicesCharm.Visibility = Visibility.Hidden;
            SettingsCharm.Visibility = Visibility.Hidden;

            SearchCharmInactive.Visibility = Visibility.Visible;
            ShareCharmInactive.Visibility = Visibility.Visible;
            NoColor.Visibility = Visibility.Visible;
            DevicesCharmInactive.Visibility = Visibility.Visible;
            SettingsCharmInactive.Visibility = Visibility.Visible;
        }

        private void Devices_MouseUp(object sender, MouseButtonEventArgs e)
        {
            /* the old Devices charm behavior chose to select between four screen options - this has been removed
            byte winKey = (byte)KeyInterop.VirtualKeyFromKey(Key.LWin);
            byte pKey = (byte)KeyInterop.VirtualKeyFromKey(Key.P);
            const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
            const uint KEYEVENTF_KEYUP = 0x0002;

            if (this.IsActive == true)
            {
                _ = keybd_event(winKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
                _ = keybd_event(pKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
                _ = keybd_event(winKey, 0, KEYEVENTF_KEYUP, 0);
                _ = keybd_event(pKey, 0, KEYEVENTF_KEYUP, 0);
            }
            */
            findDevices = true;

            swipeIn = false;
            keyboardShortcut = false;
            charmsAppear = false;
            charmsUse = false;
            charmsActivate = false;
            pokeCharms = false;

            if (useAnimations == false)
            {
                this.Opacity = 0.000;
                CharmsClock.Opacity = 0.000;

                var brush = (Brush)converter.ConvertFromString("#00111111");
                Background = brush;
            }
            mouseIn = false;

            SearchDown.Visibility = Visibility.Hidden;
            ShareDown.Visibility = Visibility.Hidden;
            WinDown.Visibility = Visibility.Hidden;
            DevicesDown.Visibility = Visibility.Hidden;
            SettingsDown.Visibility = Visibility.Hidden;

            SearchText.Visibility = Visibility.Hidden;
            ShareText.Visibility = Visibility.Hidden;
            WinText.Visibility = Visibility.Hidden;
            DevicesText.Visibility = Visibility.Hidden;
            SettingsText.Visibility = Visibility.Hidden;

            SearchCharm.Visibility = Visibility.Hidden;
            ShareCharm.Visibility = Visibility.Hidden;
            MetroColor.Visibility = Visibility.Hidden;
            DevicesCharm.Visibility = Visibility.Hidden;
            SettingsCharm.Visibility = Visibility.Hidden;

            SearchCharmInactive.Visibility = Visibility.Visible;
            ShareCharmInactive.Visibility = Visibility.Visible;
            NoColor.Visibility = Visibility.Visible;
            DevicesCharmInactive.Visibility = Visibility.Visible;
            SettingsCharmInactive.Visibility = Visibility.Visible;
        }

        private async void Settings_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (isMetro == true)
            {
                RECT rct = new RECT();
                GetWindowRect(mWnd, ref rct);
                SetCursorPos(rctLeft + 21, rctTop + 14);

                //perform click            
                mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                await Task.Delay(333);
                mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                await Task.Delay(333);
                byte tKey = (byte)KeyInterop.VirtualKeyFromKey(Key.T);
                const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
                const uint KEYEVENTF_KEYUP = 0x0002;
                _ = keybd_event(tKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
                await Task.Delay(100);
                _ = keybd_event(tKey, 0, KEYEVENTF_KEYUP, 0);
            }
            else
            {
                openSettings = true;
                swipeIn = false;
                keyboardShortcut = false;
                charmsAppear = false;
                charmsUse = false;
                charmsActivate = false;
                pokeCharms = false;
                if (useAnimations == false)
                {
                    this.Opacity = 0.000;
                    CharmsClock.Opacity = 0.000;

                    var brush = (Brush)converter.ConvertFromString("#00111111");
                    Background = brush;
                }
                mouseIn = false;

                SearchDown.Visibility = Visibility.Hidden;
                ShareDown.Visibility = Visibility.Hidden;
                WinDown.Visibility = Visibility.Hidden;
                DevicesDown.Visibility = Visibility.Hidden;
                SettingsDown.Visibility = Visibility.Hidden;

                SearchText.Visibility = Visibility.Hidden;
                ShareText.Visibility = Visibility.Hidden;
                WinText.Visibility = Visibility.Hidden;
                DevicesText.Visibility = Visibility.Hidden;
                SettingsText.Visibility = Visibility.Hidden;

                SearchCharm.Visibility = Visibility.Hidden;
                ShareCharm.Visibility = Visibility.Hidden;
                MetroColor.Visibility = Visibility.Hidden;
                DevicesCharm.Visibility = Visibility.Hidden;
                SettingsCharm.Visibility = Visibility.Hidden;

                SearchCharmInactive.Visibility = Visibility.Visible;
                ShareCharmInactive.Visibility = Visibility.Visible;
                NoColor.Visibility = Visibility.Visible;
                DevicesCharmInactive.Visibility = Visibility.Visible;
                SettingsCharmInactive.Visibility = Visibility.Visible;
            }
        }
        private System.Timers.Timer t = null;
        private readonly Dispatcher dispatcher = Dispatcher.CurrentDispatcher;
        private void _initTimer()
        {
            t = new System.Timers.Timer();
            t.Interval = 15;
            t.Elapsed += OnTimedEvent;
            t.AutoReset = true;
            t.Enabled = true;
            t.Start();
        }

        //charms bar begin
        private void OnTimedEvent(object sender, ElapsedEventArgs e)
        {
            dispatcher.BeginInvoke((Action)(() =>
            {
                if (this.IsActive == false && charmsAppear == false)
                {
                    IntPtr handle = GetForegroundWindow();
                    if (isMetroApp(handle) == true)
                    {
                        ActiveWindow.Content = handle.ToString();
                        charmsDelay = charmsDelay2;
                        isMetro = true;
                        isGui = false;
                    }
                else if (isWinGui(handle) == true)
                {
                    ActiveWindow.Content = handle.ToString();
                    charmsDelay = charmsDelay2;
                    isMetro = false;
                    isGui = true;
                }
                else
                {
                    ActiveWindow.Content = handle.ToString();
                    charmsDelay = myCharmsDelay;
                    isMetro = false;
                    isGui = false;
                }
                }

                if (this.IsActive == true)
                {
                    IntPtr handle = GetForegroundWindow();
                    ActiveCharm.Content = handle.ToString();
                }

                if (IHOb > 0.001 && IHOb < 0.092)
                {
                    SetForegroundWindow(Int32.Parse(ActiveWindow.Content.ToString()));
                }

                try
                {
                    RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\ImmersiveShell\\EdgeUi", false);
                    if (key != null)
                    {
                        // get value
                        string m = key.GetValue("DisableTRCorner", -1, RegistryValueOptions.None).ToString();
                        string n = key.GetValue("DisableBRCorner", -1, RegistryValueOptions.None).ToString();
                        string customDelay = key.GetValue("CharmsBarDesktopDelay", -1, RegistryValueOptions.None).ToString();
                        string customDelay2 = key.GetValue("CharmsBarImmersiveDelay", -1, RegistryValueOptions.None).ToString();
                        string noCharms = key.GetValue("DisableCharmsHint", -1, RegistryValueOptions.None).ToString();
                        string noSearch = key.GetValue("DisableSearchCharm", -1, RegistryValueOptions.None).ToString(); //this is not in Windows 8.1
                        string noShare = key.GetValue("DisableShareCharm", -1, RegistryValueOptions.None).ToString(); //this is not in Windows 8.1
                        string noWin = key.GetValue("DisableStartCharm", -1, RegistryValueOptions.None).ToString(); //this is not in Windows 8.1
                        string noDevices = key.GetValue("DisableDevicesCharm", -1, RegistryValueOptions.None).ToString(); //this is not in Windows 8.1
                        string noSettings = key.GetValue("DisableSettingsCharm", -1, RegistryValueOptions.None).ToString(); //this is not in Windows 8.1, but used to replicate the Windows 10 beta design WITHOUT Settings
                        string noClock = key.GetValue("DisableCharmsClock", -1, RegistryValueOptions.None).ToString(); //this is not in Windows 8.1, but used to remove the Charms Clock
                        string noFade = key.GetValue("DisableStartFade", -1, RegistryValueOptions.None).ToString(); //in Windows 8.1 animations must be disabled to do the trick, but here it's used to remove the fading effect of the start charm whilst keeping the Charms Bar animations.
                        string charmMenuUse = key.GetValue("EnableCharmsMenu", -1, RegistryValueOptions.None).ToString(); //this is not in Windows 8.1, but used to remove the Charms Clock

                            if (noClock == "-1")
                            {
                                noClocks.Content = "0";
                            }
                            else
                            {
                                noClocks.Content = noClock;
                            }

                        if (charmMenuUse == "-1")
                        {
                            useMenu.Content = "0";
                        }
                        else
                        {
                            useMenu.Content = charmMenuUse;
                        }

                            if (m == "-1")
                            {
                                vn.Content = "0";
                            }
                            else
                            {
                                vn.Content = m;

                            }
                            if (n == "-1")
                            {
                                vn2.Content = "0";
                            }
                            else
                            {
                                vn2.Content = n;
                            }

                        if (IHOb < 0.02 && charmsUse == false)
                        {
                            if (noCharms != "0" && noCharms != "-1")
                            {
                                NoColor.Visibility = Visibility.Hidden;
                                SearchCharmInactive.Visibility = Visibility.Hidden;
                                ShareCharmInactive.Visibility = Visibility.Hidden;
                                WinCharmInactive.Visibility = Visibility.Hidden;
                                DevicesCharmInactive.Visibility = Visibility.Hidden;
                                SettingsCharmInactive.Visibility = Visibility.Hidden;
                            }
                            else
                            {
                                NoColor.Visibility = Visibility.Visible;
                                SearchCharmInactive.Visibility = Visibility.Visible;
                                ShareCharmInactive.Visibility = Visibility.Visible;
                                WinCharmInactive.Visibility = Visibility.Visible;

                                if (vn4.Content != "0" && vn4.Content != "-1")
                                {
                                    DevicesCharmInactive.Visibility = Visibility.Hidden;
                                }
                                else
                                {
                                    DevicesCharmInactive.Visibility = Visibility.Visible;
                                }

                                if (vn3.Content != "0" && vn3.Content != "-1")
                                {
                                    SettingsCharmInactive.Visibility = Visibility.Hidden;
                                }
                                else
                                {
                                    SettingsCharmInactive.Visibility = Visibility.Visible;
                                }
                            }

                            if (noSettings == "-1")
                            {
                                vn3.Content = "0";
                            }
                            else
                            {
                                vn3.Content = noSettings;
                            }

                            if (noDevices == "-1")
                            {
                                vn4.Content = "0";
                            }
                            else
                            {
                                vn4.Content = noDevices;
                            }

                            if (noSearch == "-1")
                            {
                                vn5.Content = "0";
                            }
                            else
                            {
                                vn5.Content = noSearch;
                            }

                            if (noShare == "-1")
                            {
                                vn6.Content = "0";
                            }
                            else
                            {
                                vn6.Content = noShare;
                            }

                            if (noWin == "-1")
                            {
                                vn7.Content = "0";
                            }
                            else
                            {
                                vn7.Content = noWin;
                            }

                            if (noFade == "-1")
                            {
                                noFades.Content = "0";
                            }
                            else
                            {
                                noFades.Content = noFade;
                            }

                            if (customDelay == "-1" || customDelay == null)
                            {
                                myCharmsDelay = 100;
                            }
                            else
                            {
                                myCharmsDelay = Int32.Parse(customDelay);
                            }

                            if (customDelay2 == "-1" || customDelay2 == null)
                            {
                                charmsDelay2 = 100;
                            }
                            else
                            {
                                charmsDelay2 = Int32.Parse(customDelay2);
                            }
                        }
                        key.Close();
                    }
                    else
                    {
                        string m = "0";
                        string n = "0";
                        string customDelay = "-1";
                        string customDelay2 = "-1";
                        string noCharms = "0";
                        string noSearch = "0";
                        string noShare = "0";
                        string noWin = "0";
                        string noDevices = "0";
                        string noSettings = "0";
                        string noClock = "0";
                        string noFade = "0";
                        string charmMenuUse = "0";

                        if (charmMenuUse == "-1")
                        {
                            useMenu.Content = "0";
                        }
                        else
                        {
                            useMenu.Content = charmMenuUse;
                        }

                        if (IHOb < 0.02 && charmsUse == false)
                        {
                            if (noCharms != "0" && noCharms != "-1")
                            {
                                NoColor.Visibility = Visibility.Hidden;
                                SearchCharmInactive.Visibility = Visibility.Hidden;
                                ShareCharmInactive.Visibility = Visibility.Hidden;
                                WinCharmInactive.Visibility = Visibility.Hidden;
                                DevicesCharmInactive.Visibility = Visibility.Hidden;
                                SettingsCharmInactive.Visibility = Visibility.Hidden;
                            }
                            else
                            {
                                NoColor.Visibility = Visibility.Visible;
                                SearchCharmInactive.Visibility = Visibility.Visible;
                                ShareCharmInactive.Visibility = Visibility.Visible;
                                WinCharmInactive.Visibility = Visibility.Visible;

                                if (vn4.Content != "0" && vn4.Content != "-1")
                                {
                                    DevicesCharmInactive.Visibility = Visibility.Hidden;
                                }
                                else
                                {
                                    DevicesCharmInactive.Visibility = Visibility.Visible;
                                }

                                if (vn3.Content != "0" && vn3.Content != "-1")
                                {
                                    SettingsCharmInactive.Visibility = Visibility.Hidden;
                                }
                                else
                                {
                                    SettingsCharmInactive.Visibility = Visibility.Visible;
                                }
                            }

                            if (noSettings == "-1")
                            {
                                vn3.Content = "0";
                            }
                            else
                            {
                                vn3.Content = noSettings;
                            }

                            if (noDevices == "-1")
                            {
                                vn4.Content = "0";
                            }
                            else
                            {
                                vn4.Content = noDevices;
                            }

                            if (noSearch == "-1")
                            {
                                vn5.Content = "0";
                            }
                            else
                            {
                                vn5.Content = noSearch;
                            }

                            if (noShare == "-1")
                            {
                                vn6.Content = "0";
                            }
                            else
                            {
                                vn6.Content = noShare;
                            }

                            if (noWin == "-1")
                            {
                                vn7.Content = "0";
                            }
                            else
                            {
                                vn7.Content = noWin;
                            }

                            if (noClock == "-1")
                            {
                                noClocks.Content = "0";
                            }
                            else
                            {
                                noClocks.Content = noClock;
                            }

                            if (noFade == "-1")
                            {
                                noFades.Content = "0";
                            }
                            else
                            {
                                noFades.Content = noFade;
                            }
                            myCharmsDelay = 100;
                            charmsDelay2 = 100;
                        }
                    }
                }
                catch (Exception ex)  //just for demonstration...it's always best to handle specific exceptions
                {
                    //react appropriately
                }

                if (System.Globalization.CultureInfo.CurrentCulture.Name == "jp")
                {
                    SearchText.Content = "  捜索";
                    ShareText.Content = "共有";
                    WinText.Content = "スタート";
                    DevicesText.Content = "デバイス";
                    SettingsText.Content = "設定";
                }

                if (System.Globalization.CultureInfo.CurrentCulture.Name == "de")
                {
                    SearchText.Content = "Suchen";
                    ShareText.Content = "Freigeben";
                    WinText.Content = "Start";
                    DevicesText.Content = "Geräte";
                    SettingsText.Content = "Einstellungen";
                }

                if (findDevices == true || openSettings == true)
                {
                    findTimer += 1;
                }

                if (findTimer > 50)
                {
                    // Prepare the process to run
                    ProcessStartInfo start = new ProcessStartInfo();
                    if (findDevices == true)
                    {
                        start.Arguments = "ms-settings-connectabledevices:devicediscovery";
                    }

                    if (openSettings == true)
                    {
                        start.Arguments = "ms-settings:";
                    }
                    // Enter the executable to run, including the complete path
                    start.FileName = "explorer";
                    // Do you want to show a console window?
                    start.WindowStyle = ProcessWindowStyle.Hidden;
                    start.CreateNoWindow = true;
                    int exitCode;
                    // Run the external process & wait for it to finish
                    using (Process proc = Process.Start(start))
                    {
                        proc.WaitForExit();

                        // Retrieve the app's exit code
                        exitCode = proc.ExitCode;
                    }
                    findDevices = false;
                    openSettings = false;
                    findTimer = 0;
                }

                if (vn3.Content != "0" && vn3.Content != "-1")
                {
                    TechnicalPreview.Height = new GridLength(0, GridUnitType.Pixel);
                    SettingsBG.Visibility = Visibility.Hidden;
                }
                else
                {
                    TechnicalPreview.Height = new GridLength(75, GridUnitType.Pixel);
                    SettingsBG.Visibility = Visibility.Visible;
                }

                if (vn4.Content != "0" && vn4.Content != "-1")
                {
                    DeveloperPreview.Height = new GridLength(0, GridUnitType.Pixel);
                    DevicesBG.Visibility = Visibility.Hidden;
                }
                else
                {
                    DeveloperPreview.Height = new GridLength(75, GridUnitType.Pixel);
                    DevicesBG.Visibility = Visibility.Visible;
                }

                if (vn5.Content != "0" && vn5.Content != "-1")
                {
                    ReleasePreview.Height = new GridLength(0, GridUnitType.Pixel);
                    SearchBG.Visibility = Visibility.Hidden;
                }
                else
                {
                    ReleasePreview.Height = new GridLength(75, GridUnitType.Pixel);
                    SearchBG.Visibility = Visibility.Visible;
                }

                if (vn6.Content != "0" && vn6.Content != "-1")
                {
                    RTM.Height = new GridLength(0, GridUnitType.Pixel);
                    ShareBG.Visibility = Visibility.Hidden;
                }
                else
                {
                    RTM.Height = new GridLength(75, GridUnitType.Pixel);
                    ShareBG.Visibility = Visibility.Visible;
                }

                if (vn7.Content != "0" && vn7.Content != "-1")
                {
                    WinBG.Visibility = Visibility.Hidden;
                }
                else
                {
                    WinBG.Visibility = Visibility.Visible;
                }

                int mainwidth = 0;
                int mainheight = 0;
                int twowidth = 0;
                int twoheight = 0;
                int twoX = 0;
                int threewidth = 0;
                int threeheight = 0;
                int threeX = 0;
                int fourwidth = 0;
                int fourheight = 0;
                int fourX = 0;
                int fivewidth = 0;
                int fiveheight = 0;
                int fiveX = 0;
                int sixwidth = 0;
                int sixheight = 0;
                int sixX = 0;
                int sevenwidth = 0;
                int sevenheight = 0;
                int sevenX = 0;
                int eightwidth = 0;
                int eightheight = 0;
                int eightX = 0;
                int ninewidth = 0;
                int nineheight = 0;
                int nineX = 0;
                int tenwidth = 0;
                int tenheight = 0;
                int tenX = 0;
                int elevenwidth = 0;
                int elevenheight = 0;
                int elevenX = 0;

                Mouse.Capture(this, CaptureMode.SubTree);
                Point pointToWindow = Mouse.GetPosition(this);
                Point pointToScreen = PointToScreen(pointToWindow);
                Mouse.Capture(null, CaptureMode.SubTree);


                Grid.SetRow(SearchBG, 3);
                Grid.SetRow(ShareBG, 4);
                Grid.SetRow(WinBG, 6);
                Grid.SetRow(DevicesBG, 7);
                if (vn4.Content != "0" && vn4.Content != "-1")
                {
                    Grid.SetRow(SettingsBG, 7);
                }
                else
                {
                    Grid.SetRow(SettingsBG, 8);
                }

                int numVal = Int32.Parse(pointToScreen.X.ToString());
                int numVal2 = Int32.Parse(pointToScreen.Y.ToString());

                if (useMenu.Content != "0")
                {
                    if (charmsMenuOpen == true)
                    {
                        IHOb = 1.00;
                        var dispWidth = SystemParameters.PrimaryScreenWidth;
                        var dispHeight = SystemParameters.PrimaryScreenHeight;
                        CharmsMenu.Top = dispHeight - 200;
                        CharmsMenu.Opacity = IHOb;
                        CharmsClock.Opacity = IHOb;
                        CharmsClock.Left = dispWidth - 527;
                        CharmsClock.Show();
                        this.Left = -90;
                    }

                    if (charmsMenuOpen == false && numVal < 136 && numVal2 > Screen.AllScreens[0].Bounds.Height - 2)
                    {
                        CharmsMenu.Show();
                        charmsMenuOpen = true;
                    }

                    if (numVal > 136 && numVal2 < Screen.AllScreens[0].Bounds.Height - 2)
                    {
                        CharmsMenu.Hide();
                        charmsMenuOpen = false;
                    }
                }
                try
                {
                    for (int index = 0; index < Screen.AllScreens.Length;)
                    {
                        if (index == 0)
                        {
                            //mainwidth = Screen.AllScreens[0].Bounds.Width;
                            //mainheight = Screen.AllScreens[0].Bounds.Height;
                            //mainX = Screen.AllScreens[0].Bounds.Location.X;

                            mainwidth = Screen.PrimaryScreen.Bounds.Width;
                            mainheight = Screen.PrimaryScreen.Bounds.Height;
                            mainX = Screen.PrimaryScreen.Bounds.Location.X;
                        }
                        if (index == 1)
                        {
                            //two monitors
                            twowidth = Screen.AllScreens[1].Bounds.Width;
                            twoheight = Screen.AllScreens[1].Bounds.Height;
                            twoX = Screen.AllScreens[1].Bounds.Location.X;
                        }

                        if (index == 2)
                        {
                            //three monitors
                            threewidth = Screen.AllScreens[2].Bounds.Width;
                            threeheight = Screen.AllScreens[2].Bounds.Height;
                            threeX = Screen.AllScreens[2].Bounds.Location.X;
                        }

                        if (index == 3)
                        {
                            //four monitors
                            fourwidth = Screen.AllScreens[3].Bounds.Width;
                            fourheight = Screen.AllScreens[3].Bounds.Height;
                            fourX = Screen.AllScreens[3].Bounds.Location.X;
                        }

                        if (index == 4)
                        {
                            //five monitors
                            fivewidth = Screen.AllScreens[4].Bounds.Width;
                            fiveheight = Screen.AllScreens[4].Bounds.Height;
                            fiveX = Screen.AllScreens[4].Bounds.Location.X;
                        }

                        if (index == 5)
                        {
                            //six monitors
                            sixwidth = Screen.AllScreens[5].Bounds.Width;
                            sixheight = Screen.AllScreens[5].Bounds.Height;
                            sixX = Screen.AllScreens[5].Bounds.Location.X;
                        }

                        if (index == 6)
                        {
                            //seven monitors
                            sevenwidth = Screen.AllScreens[6].Bounds.Width;
                            sevenheight = Screen.AllScreens[6].Bounds.Height;
                            sevenX = Screen.AllScreens[6].Bounds.Location.X;
                        }

                        if (index == 7)
                        {
                            //eight monitors
                            eightwidth = Screen.AllScreens[7].Bounds.Width;
                            eightheight = Screen.AllScreens[7].Bounds.Height;
                            eightX = Screen.AllScreens[7].Bounds.Location.X;

                        }

                        if (index == 8)
                        {
                            //nine monitors
                            ninewidth = Screen.AllScreens[8].Bounds.Width;
                            nineheight = Screen.AllScreens[8].Bounds.Height;
                            nineX = Screen.AllScreens[8].Bounds.Location.X;
                        }

                        if (index == 9)
                        {
                            //ten monitors
                            tenwidth = Screen.AllScreens[9].Bounds.Width;
                            tenheight = Screen.AllScreens[9].Bounds.Height;
                            tenX = Screen.AllScreens[9].Bounds.Location.X;
                        }

                        if (index == 10)
                        {
                            //eleven monitors
                            elevenwidth = Screen.AllScreens[10].Bounds.Width;
                            elevenheight = Screen.AllScreens[10].Bounds.Height;
                            elevenX = Screen.AllScreens[10].Bounds.Location.X;
                        }

                        if (index == 11)
                        {
                        }
                        index++;
                    }
                }

                catch (Exception err)
                {
                    mainwidth = Screen.AllScreens[0].Bounds.Width;
                    mainheight = Screen.AllScreens[0].Bounds.Height;
                    twowidth = 0;
                    twoheight = 0;
                    twoX = 0;
                    threewidth = 0;
                    threeheight = 0;
                    threeX = 0;
                    fourwidth = 0;
                    fourheight = 0;
                    fourX = 0;
                    fivewidth = 0;
                    fiveheight = 0;
                    fiveX = 0;
                    sixwidth = 0;
                    sixheight = 0;
                    sixX = 0;
                    sevenwidth = 0;
                    sevenheight = 0;
                    sevenX = 0;
                    eightwidth = 0;
                    eightheight = 0;
                    eightX = 0;
                    ninewidth = 0;
                    nineheight = 0;
                    nineX = 0;
                    tenwidth = 0;
                    tenheight = 0;
                    tenX = 0;
                    elevenwidth = 0;
                    elevenheight = 0;
                    elevenX = 0;
                }

                //                if (IHOb < 0.012)
                if (charmsMenuOpen == false)
                {
                    if (activeScreen == 0)
                    {
                        if (charmsUse == false)
                        {
                            this.Left = mainwidth - 86;
                        }

                        if (charmsUse == true && keyboardShortcut == false && charmsFade == true)
                        {
                            this.Left = mainwidth - 86;
                        }

                        if (mainheight > 600 - 1)
                        {
                            this.Height = mainheight;
                            this.Top = 0;
                        }
                        else
                        {
                            this.Height = 900;
                            this.Top = 0 - mainheight + 445;
                        }
                        CharmsClock.Left = 51;
                        CharmsClock.Top = mainheight - 188;
                    }
                }

                if (activeScreen == 0)
                {
                    screenwidth = mainwidth;
                    screenheight = mainheight;
                    screenX = 0;
                }

                if (activeScreen == 1)
                {
                    screenwidth = mainwidth + twowidth;
                    screenheight = twoheight;
                    screenX = twoX;
                }

                if (activeScreen == 2)
                {
                    screenwidth = mainwidth + twowidth + threewidth;
                    screenheight = threeheight;
                    screenX = threeX;
                }

                if (activeScreen == 3)
                {
                    screenwidth = mainwidth + twowidth + threewidth + fourwidth;
                    screenheight = fourheight;
                    screenX = fourX;
                }

                if (activeScreen == 4)
                {
                    screenwidth = mainwidth + twowidth + threewidth + fourwidth + fivewidth;
                    screenheight = fiveheight;
                    screenX = fiveX;
                }

                if (activeScreen == 5)
                {
                    screenwidth = mainwidth + twowidth + threewidth + fourwidth + fivewidth + sixwidth;
                    screenheight = sixheight;
                    screenX = sixX;
                }

                if (activeScreen == 6)
                {
                    screenwidth = mainwidth + twowidth + threewidth + fourwidth + fivewidth + sixwidth + sevenwidth;
                    screenheight = sevenheight;
                    screenX = sevenX;
                }

                if (activeScreen == 7)
                {
                    screenwidth = mainwidth + twowidth + threewidth + fourwidth + fivewidth + sixwidth + sevenwidth + eightwidth;
                    screenheight = eightheight;
                    screenX = eightX;
                }

                if (activeScreen == 8)
                {
                    screenwidth = mainwidth + twowidth + threewidth + fourwidth + fivewidth + sixwidth + sevenwidth + eightwidth + ninewidth;
                    screenheight = nineheight;
                    screenX = nineX;
                }

                if (activeScreen == 9)
                {
                    screenwidth = mainwidth + twowidth + threewidth + fourwidth + fivewidth + sixwidth + sevenwidth + eightwidth + ninewidth + tenwidth;
                    screenheight = tenheight;
                    screenX = tenX;
                }

                if (activeScreen == 10)
                {
                    screenwidth = mainwidth + twowidth + threewidth + fourwidth + fivewidth + sixwidth + sevenwidth + eightwidth + ninewidth + tenwidth + elevenwidth;
                    screenheight = elevenheight;
                    screenX = elevenX;
                }

                if (numVal > mainwidth + 55)
                {
                    if (swipeIn == false && activeScreen != 1)
                    {
                        activeScreen = 1;
                        screenheight = twoheight;
                    }

                    if (swipeIn == true && activeScreen != 1)
                    {
                        forceClose = true;
                        activeScreen = 1;
                        screenheight = twoheight;
                    }
                }

                if (numVal < mainwidth + 55)
                {
                    if (swipeIn == false && activeScreen != 0)
                    {
                        activeScreen = 0;
                        screenwidth = 0;
                        screenheight = 0;
                    }

                    if (swipeIn == true && activeScreen != 0)
                    {
                        forceClose = true;
                        activeScreen = 0;
                        screenwidth = 0;
                        screenheight = 0;
                    }
                }

                if (numVal > mainwidth + twowidth + 55)
                {
                    if (swipeIn == false && activeScreen != 2)
                    {
                        activeScreen = 2;
                        screenheight = threeheight;
                    }

                    if (swipeIn == true && activeScreen != 2)
                    {
                        forceClose = true;
                        activeScreen = 2;
                        screenheight = threeheight;
                    }
                }

                if (numVal > mainwidth + twowidth + threewidth + 55)
                {
                    if (swipeIn == false && activeScreen != 3)
                    {
                        activeScreen = 3;
                    }

                    if (swipeIn == true && activeScreen != 3)
                    {
                        forceClose = true;
                        activeScreen = 3;
                    }
                }

                if (numVal > mainwidth + twowidth + threewidth + fourwidth + 55)
                {
                    if (swipeIn == false && activeScreen != 4)
                    {
                        activeScreen = 4;
                    }

                    if (swipeIn == true && activeScreen != 4)
                    {
                        forceClose = true;
                        activeScreen = 4;
                    }
                }

                if (numVal > mainwidth + twowidth + threewidth + fourwidth + fivewidth + 55)
                {
                    if (swipeIn == false && activeScreen != 5)
                    {
                        activeScreen = 5;
                    }

                    if (swipeIn == true && activeScreen != 5)
                    {
                        forceClose = true;
                        activeScreen = 5;
                    }
                }

                if (numVal > mainwidth + twowidth + threewidth + fourwidth + fivewidth + sixwidth + 55)
                {
                    if (swipeIn == false && activeScreen != 6)
                    {
                        activeScreen = 6;
                    }

                    if (swipeIn == true && activeScreen != 6)
                    {
                        forceClose = true;
                        activeScreen = 6;
                    }
                }

                if (numVal > mainwidth + twowidth + threewidth + fourwidth + fivewidth + sixwidth + sevenwidth + 55)
                {
                    if (swipeIn == false && activeScreen != 7)
                    {
                        activeScreen = 7;
                    }

                    if (swipeIn == true && activeScreen != 7)
                    {
                        forceClose = true;
                        activeScreen = 7;
                    }
                }

                if (numVal > mainwidth + twowidth + threewidth + fourwidth + fivewidth + sixwidth + sevenwidth + eightwidth + 55)
                {
                    if (swipeIn == false && activeScreen != 8)
                    {
                        activeScreen = 8;
                    }

                    if (swipeIn == true && activeScreen != 8)
                    {
                        forceClose = true;
                        activeScreen = 8;
                    }
                }

                if (numVal > mainwidth + twowidth + threewidth + fourwidth + fivewidth + sixwidth + sevenwidth + eightwidth + ninewidth + 55)
                {
                    if (swipeIn == false && activeScreen != 9)
                    {
                        activeScreen = 9;
                    }

                    if (swipeIn == true && activeScreen != 9)
                    {
                        forceClose = true;
                        activeScreen = 9;
                    }
                }

                if (numVal > mainwidth + twowidth + threewidth + fourwidth + fivewidth + sixwidth + sevenwidth + eightwidth + ninewidth + tenwidth + 55)
                {
                    if (swipeIn == false && activeScreen != 10)
                    {
                        activeScreen = 10;
                    }

                    if (swipeIn == true && activeScreen != 10)
                    {
                        forceClose = true;
                        activeScreen = 10;
                    }
                }

                if (numVal > mainwidth + twowidth + threewidth + fourwidth + fivewidth + sixwidth + sevenwidth + eightwidth + ninewidth + tenwidth + elevenwidth + 55)
                {
                    if (swipeIn == false && activeScreen != 11)
                    {
                        activeScreen = 11;
                    }

                    if (swipeIn == true && activeScreen != 11)
                    {
                        forceClose = true;
                        activeScreen = 11;
                    }
                }
                try
                {
                    if (activeScreen != 0 && IHOb < 0.012)
                    {
                        if (charmsUse == false)
                        {
                            this.Left = screenwidth - 86;
                        }

                        if (charmsUse == true && keyboardShortcut == false)
                        {
                            this.Left = screenwidth - 86;
                        }
                        if (screenheight > 600 - 1)
                        {
                            this.Height = Screen.AllScreens[activeScreen].Bounds.Height;
                            this.Top = Screen.AllScreens[activeScreen].Bounds.Top;
                        }
                        else
                        {
                            this.Height = 900;
                            this.Top = 0 - Screen.AllScreens[activeScreen].Bounds.Height + 395;
                        }

                        CharmsClock.Left = Screen.AllScreens[activeScreen].Bounds.Left + 51;
                        CharmsClock.Top = Screen.AllScreens[activeScreen].Bounds.Bottom - 188;
                    }
                }

                catch (Exception err)
                {
                    if (activeScreen != 0 && IHOb < 0.012)
                    {
                        if (charmsUse == false)
                        {
                            this.Left = screenwidth - 86;
                        }

                        if (charmsUse == true && keyboardShortcut == false)
                        {
                            this.Left = screenwidth - 86;
                        }
                        if (screenheight > 600)
                        {
                            this.Height = screenheight;
                            this.Top = 0;
                        }
                        else
                        {
                            this.Height = screenheight;
                            this.Top = 0 - screenheight + 395;
                        }

                        CharmsClock.Left = screenX + 51;
                        CharmsClock.Top = screenheight - 188;
                    }
                }
                try
                {
                    bool animationsEnabled;
                    SystemParametersInfo(SPI_GETCLIENTAREAANIMATION, 0x00, out animationsEnabled, 0x00);

                    if (animationsEnabled)
                    {
                        if (useAnimations == false)
                        {
                            useAnimations = true; //Opens the charms bar with animations enabled!
                        }
                    }
                    else
                    {
                        if (useAnimations == true)
                        {
                            useAnimations = false;
                        }
                    }
                }
                catch (Win32Exception ex)
                {
                    //error
                }

                try
                {
                    if (System.Windows.Forms.Cursor.Current.IsVisible() == false)
                    {
                        usingTouch = true;
                    }
                    else
                    {
                        usingTouch = false;
                    }
                }

                catch (Exception err)
                {

                }

                if (keyboardShortcut == false && swipeIn == true && vn.Content == "1" && vn2.Content == "1")
                {
                    swipeIn = false;
                }

                if (holder == true && usingTouch == false)
                {
                    cursorStay += 1;
                }

                double charmsExpire = 140 * 2.21;
                if (holder == true && cursorStay > charmsExpire)
                {
                    pokeCharms = true;
                    outofTime = false;
                    blockRepeating = 1;
                    cursorStay = 0;
                }

                if (holder == false && cursorStay > charmsExpire)
                {
                    pokeCharms = false;
                    outofTime = false;
                    blockRepeating = 1;
                    cursorStay = 0;
                }

                if (keyboardShortcut == false)
                {
                    winStretch = 80.31;
                }

                if (keyboardShortcut == true && scrollWin != 0)
                {
                    winStretch = 140.31;
                }

                if (System.Windows.Forms.Control.MouseButtons == MouseButtons.None)
                {
                    searchActive = false;
                    shareActive = false;
                    winActive = false;
                    devicesActive = false;
                    settingsActive = false;
                }

                if (System.Windows.Forms.Control.MouseButtons == MouseButtons.Left && searchHover == true)
                {
                    searchActive = true;
                }
                else
                {
                    searchActive = false;
                }

                if (System.Windows.Forms.Control.MouseButtons == MouseButtons.Left && shareHover == true)
                {
                    shareActive = true;
                }
                else
                {
                    shareActive = false;
                }

                if (System.Windows.Forms.Control.MouseButtons == MouseButtons.Left && winHover == true)
                {
                    winActive = true;
                }
                else
                {
                    winActive = false;
                }

                if (System.Windows.Forms.Control.MouseButtons == MouseButtons.Left && devicesHover == true)
                {
                    devicesActive = true;
                }
                else
                {
                    devicesActive = false;
                }

                if (System.Windows.Forms.Control.MouseButtons == MouseButtons.Left && settingsHover == true)
                {
                    settingsActive = true;
                }
                else
                {
                    settingsActive = false;
                }

                if (SystemParameters.WindowGlassBrush.ToString() != "#FFFAFAFA")
                {
                    WinFader.Source = new BitmapImage(new Uri(@"/Assets/Images/fader.png", UriKind.Relative));
                }

                if (SystemParameters.WindowGlassBrush.ToString() == "#FFFAFAFA")
                {
                    WinFader.Source = new BitmapImage(new Uri(@"/Assets/Images/fader light.png", UriKind.Relative));
                }

                if (Keyboard.IsKeyDown(Key.LWin) && Keyboard.IsKeyDown(Key.C) && keyboardShortcut == false)
                {
                    pokeCharms = true;
                    charmsAppear = true;
                    charmsActivate = true;
                    charmsUse = true;
                    keyboardShortcut = true;
                    this.BringIntoView();
                    this.Focus();
                    this.Activate();
                }

                if (Keyboard.IsKeyDown(Key.RWin) && Keyboard.IsKeyDown(Key.C) && keyboardShortcut == false)
                {
                    pokeCharms = true;
                    charmsAppear = true;
                    charmsActivate = true;
                    charmsUse = true;
                    keyboardShortcut = true;
                    this.BringIntoView();
                    this.Focus();
                    this.Activate();
                }

                    if (charmsUse == false)
                    {
                        CharmBG.Opacity = 0.000;
                        CharmBG.Visibility = Visibility.Hidden;
                    }
                    else
                    {
                        CharmBG.Visibility = Visibility.Visible;
                    }

                //finally, a new solution to making the charms bar less hostile!
                if (preventReload == true)
                {
                    blockRepeating = 0;
                }

                if (preventReload == false)
                {
                    blockRepeating += 1;
                }

                if (outofTime == true && holder == true && cursorStay < charmsExpire)
                {
                    blockRepeating = 0;
                }

                if (charmsUse == false)
                {
                    if (swipeIn == true && charmsReturn == dasBoot && outofTime == true && blockRepeating != 0)
                    {
                        if (useAnimations == true)
                        {
                            scrollSearch = 200;
                            scrollShare = 150;
                            scrollWin = 100;
                            scrollDevices = 150;
                            scrollSettings = 200;

                            textSearch = 190;
                            textShare = 150;
                            textWin = 100;
                            textDevices = 150;
                            textSettings = 200;
                        }
                        else
                        {
                            scrollSearch = 0;
                            scrollShare = 0;
                            scrollWin = 0;
                            scrollDevices = 0;
                            scrollSettings = 0;

                            textSearch = 0;
                            textShare = 0;
                            textWin = 0;
                            textDevices = 0;
                            textSettings = 0;
                        }

                        charmsTimer = 0;
                        charmsWait = 0;
                        dasBoot += 1;
                        charmsReturn = dasBoot;
                        outofTime = false;
                        swipeIn = true;
                    }

                }

                if (charmsUse == false)
                {
                    if (useAnimations == true)
                    {
                        if (CharmBG.Opacity > 0.000 && outofTime == false && forceClose == false)
                        {
                            FadeBlocker.Opacity -= 0.1;
                            CharmBG.Opacity -= 0.1;
                            WinCharm.Opacity -= 0.1;
                            MetroColor.Opacity -= 0.1;
                        }
                        if (IHOb > 0.5)
                        {
                            dasSlide = -350;
                        }
                        else
                        {
                            dasSlide = -410;
                        }
                    }

                    if (useAnimations == false)
                    {
                        if (this.Opacity > 0.000 && outofTime == false && forceClose == false)
                        {
                            FadeBlocker.Opacity = 1.00;
                            WinCharm.Opacity = 1.00;
                            MetroColor.Opacity = 1.00;
                            this.Opacity = 1.00;
                        }
                        dasSlide = 44450;
                    }
                }

                if (charmsUse == true && useAnimations == true)
                {
                    dasSlide += 8;
                }

                if (charmsUse == true && useAnimations == false)
                {
                    dasSlide = 1945;
                }

                if (IHOb > 0.1 && IHOb < 0.9 && keyboardShortcut == false)
                {
                    charmsFade = true;
                    if (charmsUse == true && numVal < mainwidth - 116 && activeScreen == 0)
                    {
                        ignoreMouseIn = true;
                    }

                    if (charmsUse == true && numVal < screenwidth - 116 && activeScreen > 0)
                    {
                        ignoreMouseIn = true;
                    }
                }

                if (charmsUse == false && useAnimations == true && keyboardShortcut == true)
                {
                    activeIcon = 2;
                }

                if (charmsUse == false && useAnimations == true && keyboardShortcut == false)
                {
                    activeIcon = 2;
                }

                if (IHOb == 0.000 || IHOb == 1.0)
                {
                    charmsFade = false;
                }

                if (noFades.Content != "1")
                {
                    WinFader.Margin = new Thickness(dasSlide, 0, 0, 0);

                }

                if (forceClose == true && IHOb < 0.012)
                {
                    forceClose = false;
                }

                if (IHOb <= 0.012 && useAnimations == true && charmsTimer == 1 || keyboardShortcut == true && IHOb <= 0.012 && useAnimations == true)
                {
                    scrollSearch = 200;
                    scrollShare = 150;
                    scrollWin = 100;
                    scrollDevices = 150;
                    scrollSettings = 200;

                    textSearch = 190;
                    textShare = 150;
                    textWin = 100;
                    textDevices = 150;
                    textSettings = 200;
                }

                if (IHOb <= 0.000 && useAnimations == false && charmsTimer == 1 || keyboardShortcut == true && IHOb <= 0.000 && useAnimations == false)
                {
                    scrollSearch = 0;
                    scrollShare = 0;
                    scrollWin = 0;
                    scrollDevices = 0;
                    scrollSettings = 0;

                    textSearch = 0;
                    textShare = 0;
                    textWin = 0;
                    textDevices = 0;
                    textSettings = 0;
                }

                if (charmsUse == true && CharmsClock.IsVisible && keyboardShortcut == true && useAnimations == true && SystemParameters.HighContrast == false)
                {
                    var brush = (Brush)converter.ConvertFromString("#111111");
                    CharmBG.Background = brush;
                    CharmBG.Opacity = 1;
                }

                if (charmsUse == true && CharmsClock.IsVisible && keyboardShortcut == true && useAnimations == true && SystemParameters.HighContrast == true)
                {
                    CharmBG.Opacity = 1;
                }

                if (swipeIn == true && numVal < mainwidth - 116 && keyboardShortcut == false && activeScreen == 0 || swipeIn == true && numVal < screenwidth - 116 && keyboardShortcut == false && activeScreen > 0)
                {
                    dasSwiper = true;
                }
                else
                {
                    dasSwiper = false;
                }

                if (charmsUse == false && charmsAppear == false && System.Windows.Forms.Control.MouseButtons != MouseButtons.None)
                {
                    charmsTimer = 0;
                    pokeCharms = false;
                }

                if (charmsUse == false)
                {
                    SearchHover.Visibility = Visibility.Hidden;
                    ShareHover.Visibility = Visibility.Hidden;
                    WinHover.Visibility = Visibility.Hidden;
                    DevicesHover.Visibility = Visibility.Hidden;
                    SettingsHover.Visibility = Visibility.Hidden;

                    SearchDown.Visibility = Visibility.Hidden;
                    ShareDown.Visibility = Visibility.Hidden;
                    WinDown.Visibility = Visibility.Hidden;
                    DevicesDown.Visibility = Visibility.Hidden;
                    SettingsDown.Visibility = Visibility.Hidden;
                }

                if (this.ShowInTaskbar == false)
                {
                    if (charmsUse == false && IHOb < 0.012)
                    {
                        forceClose = false;
                        ignoreMouseIn = false;
                        CharmsClock.Hide();
                    }

                    if (charmsUse == true || charmsMenuOpen == true)
                    {
                        if (noClocks.Content == "-1" || noClocks.Content == "0")
                        {
                            CharmsClock.Show();
                        }
                    }

                    if (WinCharm.Visibility == Visibility.Hidden && charmsUse == false)
                    {
                        var brush = (Brush)converter.ConvertFromString("#00111111");
                        Background = brush;
                    }

                    if (useAnimations == true)
                    {
                        if (charmsAppear == true && numVal < mainwidth - 116 && keyboardShortcut == false && activeScreen == 0 || charmsAppear == true && numVal < screenwidth - 116 && keyboardShortcut == false && activeScreen > 0 || charmsAppear == true && keyboardShortcut == true && this.IsActive == false || charmsAppear == true && keyboardShortcut == true && escKey == true || ignoreMouseIn == true && pokeCharms == true || outofTime == true && pokeCharms == true || forceClose == true)
                        {
                            IHOb -= 0.141;
                            SearchCharmInactive.Opacity = IHOb;
                            ShareCharmInactive.Opacity = IHOb;
                            NoColor.Opacity = IHOb;
                            DevicesCharmInactive.Opacity = IHOb;
                            SettingsCharmInactive.Opacity = IHOb;
                        }
                    }

                    if (useAnimations == false)
                    {
                        if (charmsAppear == true && numVal < mainwidth - 116 && keyboardShortcut == false && activeScreen == 0 || charmsAppear == true && numVal < screenwidth - 116 && keyboardShortcut == false && activeScreen > 0 || charmsAppear == true && keyboardShortcut == true && this.IsActive == false || charmsAppear == true && keyboardShortcut == true && escKey == true || ignoreMouseIn == true && pokeCharms == true || outofTime == true && pokeCharms == true || forceClose == true)
                        {
                            IHOb = 0.00;
                            SearchCharmInactive.Opacity = IHOb;
                            ShareCharmInactive.Opacity = IHOb;
                            NoColor.Opacity = IHOb;
                            DevicesCharmInactive.Opacity = IHOb;
                            SettingsCharmInactive.Opacity = IHOb;
                        }
                    }

                    if (charmsAppear == true && IHOb < 0.000 && keyboardShortcut == true)
                    {
                        keyboardShortcut = false;
                    }

                    if (useAnimations == true)
                    {
                        if (charmsAppear == false && IHOb > 0.000 && keyboardShortcut == false && forceClose == false || dasSwiper == true || charmsAppear == true && escKey == true && keyboardShortcut == true || outofTime == true && pokeCharms == true && keyboardShortcut == false || forceClose == true)
                        {
                            IHOb -= 0.061;
                            this.Opacity = IHOb;
                            CharmsClock.Opacity = IHOb;
                        }
                    }

                    if (useAnimations == false)
                    {
                        if (charmsAppear == false && IHOb > 0.000 && keyboardShortcut == false && forceClose == false || dasSwiper == true || charmsAppear == true && escKey == true && keyboardShortcut == true || outofTime == true && pokeCharms == true && keyboardShortcut == false || forceClose == true)
                        {
                            IHOb = 0.00;
                            this.Opacity = IHOb;
                            CharmsClock.Opacity = IHOb;
                        }
                    }

                    if (activeScreen == 0)
                    {
                        if (charmsAppear == true && numVal < mainwidth - 116 && IHOb < 0.000 && escKey == false && keyboardShortcut == false && outofTime == false && ignoreMouseIn == false && forceClose == false)
                        {
                            IHOb = 1.0;
                            SearchCharmInactive.Opacity = IHOb;
                            ShareCharmInactive.Opacity = IHOb;
                            NoColor.Opacity = IHOb;
                            DevicesCharmInactive.Opacity = IHOb;
                            SettingsCharmInactive.Opacity = IHOb;
                            this.Opacity = 0.000;
                            charmsAppear = false;
                            charmsTimer = 0;
                            pokeCharms = false;
                        }
                    }

                    if (activeScreen > 0)
                    {
                        if (charmsAppear == true && numVal < screenwidth - 116 && IHOb < 0.000 && escKey == false && keyboardShortcut == false && outofTime == false && ignoreMouseIn == false && forceClose == false)
                        {
                            IHOb = 1.0;
                            SearchCharmInactive.Opacity = IHOb;
                            ShareCharmInactive.Opacity = IHOb;
                            NoColor.Opacity = IHOb;
                            DevicesCharmInactive.Opacity = IHOb;
                            SettingsCharmInactive.Opacity = IHOb;
                            this.Opacity = 0.000;
                            charmsAppear = false;
                            charmsTimer = 0;
                        }
                    }

                    if (IHOb < 0.000 && keyboardShortcut == false)
                    {
                        SearchDown.Visibility = Visibility.Hidden;
                        ShareDown.Visibility = Visibility.Hidden;
                        WinDown.Visibility = Visibility.Hidden;
                        DevicesDown.Visibility = Visibility.Hidden;
                        SettingsDown.Visibility = Visibility.Hidden;

                        SearchText.Visibility = Visibility.Hidden;
                        ShareText.Visibility = Visibility.Hidden;
                        WinText.Visibility = Visibility.Hidden;
                        DevicesText.Visibility = Visibility.Hidden;
                        SettingsText.Visibility = Visibility.Hidden;

                        SearchCharm.Visibility = Visibility.Hidden;
                        ShareCharm.Visibility = Visibility.Hidden;
                        MetroColor.Visibility = Visibility.Hidden;
                        DevicesCharm.Visibility = Visibility.Hidden;
                        SettingsCharm.Visibility = Visibility.Hidden;
                        ignoreMouseIn = false;
                        var brush = (Brush)converter.ConvertFromString("#00111111");
                        Background = brush;
                        swipeIn = false;
                        charmsUse = false;
                        charmsAppear = false;
                        charmsTimer = 0;
                        charmsWait = 0;
                        swipeIn = false;
                        keyboardShortcut = false;
                        pokeCharms = false;
                        CharmsClock.Opacity = 0.000;
                        CharmsClock.Hide();
                    }

                    if (useAnimations == true)
                    {
                        if (charmsAppear == true && pokeCharms == true && keyboardShortcut == false && IHOb < 1.0 && charmsUse == false && ignoreMouseIn == false && forceClose == false)
                        {
                            IHOb += 0.10;
                            SearchCharmInactive.Opacity = IHOb;
                            ShareCharmInactive.Opacity = IHOb;
                            NoColor.Opacity = IHOb;
                            DevicesCharmInactive.Opacity = IHOb;
                            SettingsCharmInactive.Opacity = IHOb;
                        }
                    }

                    if (useAnimations == false)
                    {
                        if (charmsAppear == true && pokeCharms == true && keyboardShortcut == false && IHOb < 1.0 && charmsUse == false && ignoreMouseIn == false && forceClose == false)
                        {
                            IHOb = 1.00;
                            SearchCharmInactive.Opacity = IHOb;
                            ShareCharmInactive.Opacity = IHOb;
                            NoColor.Opacity = IHOb;
                            DevicesCharmInactive.Opacity = IHOb;
                            SettingsCharmInactive.Opacity = IHOb;
                        }
                    }

                    if (charmsUse == false && IHOb < 0.000)
                    {
                        IHOb = 0.000;
                        ignoreMouseIn = false;
                        escKey = false;
                    }

                    if (useAnimations == true)
                    {
                        if (charmsAppear == true && IHOb < 1.1 && ignoreMouseIn == false)
                        {
                            IHOb += 0.05;
                            this.Opacity = IHOb;
                            CharmsClock.Opacity = IHOb;
                        }
                    }

                    if (useAnimations == false)
                    {
                        if (charmsAppear == true && IHOb < 1.1 && ignoreMouseIn == false)
                        {
                            IHOb = 1.00;
                            this.Opacity = IHOb;
                            CharmsClock.Opacity = IHOb;
                        }
                    }

                    if (charmsWait > charmsExpire && charmsUse == false && charmsReturn == dasBoot)
                    {
                        outofTime = true;
                        preventReload = false;
                    }

                    if (charmsWait < charmsExpire && numVal < mainwidth - 116)
                    {
                        outofTime = false;
                    }

                    if (this.Opacity < 0.1)
                    {
                        charmsWait = 0;
                    }

                    if (CharmBG.Opacity < 1.1 && charmsUse == true && charmsAppear == true && useAnimations == true && outofTime == false && forceClose == false)
                    {
                        if (keyboardShortcut == false)
                        {
                            WinCharm.Opacity += 0.1;
                            MetroColor.Opacity += 0.1;
                            FadeBlocker.Opacity += 0.1;
                            CharmBG.Opacity += 0.1;
                        }
                    }

                    if (CharmBG.Opacity < 1.1 && charmsUse == true && charmsAppear == true && useAnimations == false && outofTime == false && forceClose == false)
                    {
                        if (keyboardShortcut == false)
                        {
                            WinCharm.Opacity = 1.00;
                            MetroColor.Opacity = 1.00;
                            FadeBlocker.Opacity = 1.00;
                            CharmBG.Opacity = 1.00;
                        }
                    }

                    var searchDas = new Thickness(scrollSearch, -11, 12, -66);
                    SearchCharmInactive.Margin = searchDas;

                    var shareDas = new Thickness(scrollShare - 1, 14, 12, -66);
                    ShareCharmInactive.Margin = shareDas;

                    var winDas = new Thickness(scrollWin, -50, 14, 4);
                    NoColor.Margin = winDas;

                    var deviceDas = new Thickness(scrollDevices, -38, 12, -99);
                    DevicesCharmInactive.Margin = deviceDas;

                    var settingsDas = new Thickness(scrollSettings, 14, 12, -99);
                    SettingsCharmInactive.Margin = settingsDas;

                    var searchDas2 = new Thickness(scrollSearch, -11, 12, -66);
                    SearchCharm.Margin = searchDas2;

                    var shareDas2 = new Thickness(scrollShare - 1, 14, 12, -66);
                    ShareCharm.Margin = shareDas2;

                    var winDas2 = new Thickness(scrollWin, 11, 12, 4);
                    MetroColor.Margin = winDas2;

                    var deviceDas2 = new Thickness(scrollDevices, 13, 12, -10);
                    DevicesCharm.Margin = deviceDas2;

                    var settingsDas2 = new Thickness(scrollSettings, 14, 12, -99);
                    SettingsCharm.Margin = settingsDas2;

                    var searchDas3 = new Thickness(textSearch + 1, 38, 13.141, -44.89);
                    SearchText.Margin = searchDas3;

                    var shareDas3 = new Thickness(textShare - 0.005, 59, 12, 0);
                    ShareText.Margin = shareDas3;

                    var winDas3 = new Thickness(textWin + 0.003, -15, 12, 0);
                    WinText.Margin = winDas3;

                    var deviceDas3 = new Thickness(textDevices, 7, 12, 0);
                    DevicesText.Margin = deviceDas3;

                    var settingsDas3 = new Thickness(textSettings - 0.006, 59, 12, -63);
                    SettingsText.Margin = settingsDas3;

                    var searchDas4 = new Thickness(scrollSearch, -25, 0, 0);
                    SearchHover.Margin = searchDas4;

                    var shareDas4 = new Thickness(scrollShare, 0, 0, -25);
                    ShareHover.Margin = shareDas4;

                    var winDas4 = new Thickness(scrollWin, 25, 0, -50);
                    WinHover.Margin = winDas4;

                    var deviceDas4 = new Thickness(scrollDevices, 50, 0, -75);
                    DevicesHover.Margin = deviceDas4;

                    var settingsDas4 = new Thickness(scrollSettings, 75, 0, -100);
                    SettingsHover.Margin = settingsDas4;
                }

                if (this.IsActive == true && charmsUse == true)
                {

                    if (Keyboard.IsKeyUp(Key.Up) == true && Keyboard.IsKeyUp(Key.Down) == true && swipeIn == false)
                    {
                        waitTimer = 0;
                        keyboardTimer = 0;
                    }

                    if (Keyboard.IsKeyDown(Key.Up) && swipeIn == false || Keyboard.IsKeyDown(Key.Down) && swipeIn == false || Keyboard.IsKeyDown(Key.Tab) && swipeIn == false)
                    {
                        if (activeIcon == 6)
                        {
                            activeIcon = 2;
                        }
                        if (numVal < mainwidth - 116 & activeScreen == 0)
                        {
                            mouseIn = false;
                            swipeIn = false;
                        }
                        keyboardShortcut = true;
                        waitTimer += 1;
                        keyboardTimer += 1;

                        if (keyboardTimer > 5)
                        {
                            keyboardTimer = 0;
                        }
                    }

                    if (Keyboard.IsKeyDown(Key.Up) && swipeIn == false && keyboardTimer < 1 && waitTimer < 10)
                    {
                        if (activeIcon != 0)
                        {
                            activeIcon -= 1;
                        }
                        else
                        {
                            activeIcon = 4;
                        }
                        keyboardShortcut = true;
                        mouseIn = false;

                        //Search highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 0)
                        {
                            SearchHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 0)
                        {
                            SearchHover.Visibility = Visibility.Hidden;
                        }

                        //Share highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 1)
                        {
                            ShareHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 1)
                        {
                            ShareHover.Visibility = Visibility.Hidden;
                        }

                        //Start highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 2)
                        {
                            WinCharmUse = true;
                            WinHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 2)
                        {
                            WinCharmUse = false;
                            WinHover.Visibility = Visibility.Hidden;
                        }

                        //Devices highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 3)
                        {
                            DevicesHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 3)
                        {
                            DevicesHover.Visibility = Visibility.Hidden;
                        }

                        //Settings highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 4)
                        {
                            SettingsHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 4)
                        {
                            SettingsHover.Visibility = Visibility.Hidden;
                        }

                        if (mouseIn == false && keyboardShortcut == false)
                        {
                            SearchHover.Visibility = Visibility.Hidden;
                            ShareHover.Visibility = Visibility.Hidden;
                            WinCharmUse = false;
                            WinHover.Visibility = Visibility.Hidden;
                            DevicesHover.Visibility = Visibility.Hidden;
                            SettingsHover.Visibility = Visibility.Hidden;
                        }
                    }

                    if (Keyboard.IsKeyDown(Key.Down) && swipeIn == false && keyboardTimer < 1 && waitTimer < 10 && charmsUse == true || Keyboard.IsKeyDown(Key.Tab) && swipeIn == false && keyboardTimer < 1 && waitTimer < 10 && charmsUse == true)
                    {

                        if (activeIcon != 4)
                        {
                            activeIcon += 1;
                        }
                        else
                        {
                            activeIcon = 0;
                        }
                        keyboardShortcut = true;
                        mouseIn = false;

                        //Search highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 0)
                        {
                            SearchHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 0)
                        {
                            SearchHover.Visibility = Visibility.Hidden;
                        }

                        //Share highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 1)
                        {
                            ShareHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 1)
                        {
                            ShareHover.Visibility = Visibility.Hidden;
                        }

                        //Start highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 2)
                        {
                            WinCharmUse = true;
                            WinHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 2)
                        {
                            WinCharmUse = false;
                            WinHover.Visibility = Visibility.Hidden;
                        }

                        //Devices highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 3)
                        {
                            DevicesHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 3)
                        {
                            DevicesHover.Visibility = Visibility.Hidden;
                        }

                        //Settings highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 4)
                        {
                            SettingsHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 4)
                        {
                            SettingsHover.Visibility = Visibility.Hidden;
                        }

                        if (mouseIn == false && keyboardShortcut == false)
                        {
                            SearchHover.Visibility = Visibility.Hidden;
                            ShareHover.Visibility = Visibility.Hidden;
                            WinCharmUse = false;
                            WinHover.Visibility = Visibility.Hidden;
                            DevicesHover.Visibility = Visibility.Hidden;
                            SettingsHover.Visibility = Visibility.Hidden;
                        }
                    }

                    if (Keyboard.IsKeyDown(Key.Up) && swipeIn == false && keyboardTimer < 1 && waitTimer > 40 && charmsUse == true)
                    {
                        if (activeIcon != 0)
                        {
                            activeIcon -= 1;
                        }
                        else
                        {
                            activeIcon = 4;
                        }
                        keyboardShortcut = true;
                        mouseIn = false;

                        //Search highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 0)
                        {
                            SearchHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 0)
                        {
                            SearchHover.Visibility = Visibility.Hidden;
                        }

                        //Share highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 1)
                        {
                            ShareHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 1)
                        {
                            ShareHover.Visibility = Visibility.Hidden;
                        }

                        //Start highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 2)
                        {
                            WinCharmUse = true;
                            WinHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 2)
                        {
                            WinCharmUse = false;
                            WinHover.Visibility = Visibility.Hidden;
                        }

                        //Devices highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 3)
                        {
                            DevicesHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 3)
                        {
                            DevicesHover.Visibility = Visibility.Hidden;
                        }

                        //Settings highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 4)
                        {
                            SettingsHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 4)
                        {
                            SettingsHover.Visibility = Visibility.Hidden;
                        }

                        if (mouseIn == false && keyboardShortcut == false)
                        {
                            SearchHover.Visibility = Visibility.Hidden;
                            ShareHover.Visibility = Visibility.Hidden;
                            WinCharmUse = false;
                            WinHover.Visibility = Visibility.Hidden;
                            DevicesHover.Visibility = Visibility.Hidden;
                            SettingsHover.Visibility = Visibility.Hidden;
                        }
                    }

                    if (Keyboard.IsKeyDown(Key.Down) && swipeIn == false && keyboardTimer < 1 && waitTimer > 40 || Keyboard.IsKeyDown(Key.Tab) && swipeIn == false && keyboardTimer < 1 && waitTimer > 40)
                    {

                        if (activeIcon != 4)
                        {
                            activeIcon += 1;
                        }
                        else
                        {
                            activeIcon = 0;
                        }
                        keyboardShortcut = true;
                        mouseIn = false;

                        //Search highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 0)
                        {
                            SearchHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 0)
                        {
                            SearchHover.Visibility = Visibility.Hidden;
                        }

                        //Share highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 1)
                        {
                            ShareHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 1)
                        {
                            ShareHover.Visibility = Visibility.Hidden;
                        }

                        //Start highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 2)
                        {
                            WinCharmUse = true;
                            WinHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 2)
                        {
                            WinCharmUse = false;
                            WinHover.Visibility = Visibility.Hidden;
                        }

                        //Devices highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 3)
                        {
                            DevicesHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 3)
                        {
                            DevicesHover.Visibility = Visibility.Hidden;
                        }

                        //Settings highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 4)
                        {
                            SettingsHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 4)
                        {
                            SettingsHover.Visibility = Visibility.Hidden;
                        }

                        if (mouseIn == false && keyboardShortcut == false)
                        {
                            SearchHover.Visibility = Visibility.Hidden;
                            ShareHover.Visibility = Visibility.Hidden;
                            WinCharmUse = false;
                            WinHover.Visibility = Visibility.Hidden;
                            DevicesHover.Visibility = Visibility.Hidden;
                            SettingsHover.Visibility = Visibility.Hidden;
                        }
                    }
                    if (charmsUse == true)
                    {
                        if (Keyboard.IsKeyDown(Key.Enter) && keyboardShortcut == true || Keyboard.IsKeyDown(Key.Space) && keyboardShortcut == true)
                        {
                            if (activeIcon == 0)
                            {
                                byte winKey = (byte)KeyInterop.VirtualKeyFromKey(Key.LWin);
                                byte sKey = (byte)KeyInterop.VirtualKeyFromKey(Key.S);
                                const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
                                const uint KEYEVENTF_KEYUP = 0x0002;
                                _ = keybd_event(winKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
                                _ = keybd_event(sKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
                                _ = keybd_event(winKey, 0, KEYEVENTF_KEYUP, 0);
                                _ = keybd_event(sKey, 0, KEYEVENTF_KEYUP, 0);

                                swipeIn = false;
                                keyboardShortcut = false;
                                charmsAppear = false;
                                charmsUse = false;
                                charmsActivate = false;
                                pokeCharms = false;

                                if (useAnimations == false)
                                {
                                    this.Opacity = 0.000;
                                    CharmsClock.Opacity = 0.000;

                                    var brush = (Brush)converter.ConvertFromString("#00111111");
                                    Background = brush;
                                }
                                mouseIn = false;

                                SearchDown.Visibility = Visibility.Hidden;
                                ShareDown.Visibility = Visibility.Hidden;
                                WinDown.Visibility = Visibility.Hidden;
                                DevicesDown.Visibility = Visibility.Hidden;
                                SettingsDown.Visibility = Visibility.Hidden;

                                SearchText.Visibility = Visibility.Hidden;
                                ShareText.Visibility = Visibility.Hidden;
                                WinText.Visibility = Visibility.Hidden;
                                DevicesText.Visibility = Visibility.Hidden;
                                SettingsText.Visibility = Visibility.Hidden;

                                SearchCharm.Visibility = Visibility.Hidden;
                                ShareCharm.Visibility = Visibility.Hidden;
                                MetroColor.Visibility = Visibility.Hidden;
                                DevicesCharm.Visibility = Visibility.Hidden;
                                SettingsCharm.Visibility = Visibility.Hidden;

                                SearchCharmInactive.Visibility = Visibility.Visible;
                                ShareCharmInactive.Visibility = Visibility.Visible;
                                NoColor.Visibility = Visibility.Visible;
                                if (vn4.Content != "0" && vn4.Content != "-1")
                                {
                                    DevicesCharmInactive.Visibility = Visibility.Hidden;
                                }
                                else
                                {
                                    DevicesCharmInactive.Visibility = Visibility.Visible;
                                }
                                if (vn3.Content != "0" && vn3.Content != "-1")
                                {
                                    SettingsCharmInactive.Visibility = Visibility.Hidden;
                                }
                                else
                                {
                                    SettingsCharmInactive.Visibility = Visibility.Visible;
                                }
                            }

                            if (activeIcon == 1)
                            {
                                // Retrieve the window handle (HWND) of the current WinUI 3 window.
                                var hWnd = new WindowInteropHelper(this).Handle;
                                IDataTransferManagerInterop interop =
                                Windows.ApplicationModel.DataTransfer.DataTransferManager.As
                                    <IDataTransferManagerInterop>();

                                IntPtr result = interop.GetForWindow(hWnd, _dtm_iid);
                                var dataTransferManager = WinRT.MarshalInterface
                                    <Windows.ApplicationModel.DataTransfer.DataTransferManager>.FromAbi(result);

                                dataTransferManager.DataRequested += (sender, args) =>
                                {
                                    args.Request.Data.Properties.Title = " ";
                                    args.Request.Data.SetText("WinRT.Interop.WindowNative.GetWindowHandle(this)");
                                    args.Request.Data.RequestedOperation =
                                        Windows.ApplicationModel.DataTransfer.DataPackageOperation.Copy;
                                };

                                if (this.IsActive == true)
                                {
                                    interop.ShowShareUIForWindow(hWnd);
                                }

                                swipeIn = false;
                                keyboardShortcut = false;
                                charmsAppear = false;
                                charmsUse = false;
                                charmsActivate = false;
                                pokeCharms = false;
                                if (useAnimations == false)
                                {
                                    this.Opacity = 0.000;
                                    CharmsClock.Opacity = 0.000;

                                    var brush = (Brush)converter.ConvertFromString("#00111111");
                                    Background = brush;
                                }

                                mouseIn = false;

                                SearchDown.Visibility = Visibility.Hidden;
                                ShareDown.Visibility = Visibility.Hidden;
                                WinDown.Visibility = Visibility.Hidden;
                                DevicesDown.Visibility = Visibility.Hidden;
                                SettingsDown.Visibility = Visibility.Hidden;

                                SearchText.Visibility = Visibility.Hidden;
                                ShareText.Visibility = Visibility.Hidden;
                                WinText.Visibility = Visibility.Hidden;
                                DevicesText.Visibility = Visibility.Hidden;
                                SettingsText.Visibility = Visibility.Hidden;

                                SearchCharm.Visibility = Visibility.Hidden;
                                ShareCharm.Visibility = Visibility.Hidden;
                                MetroColor.Visibility = Visibility.Hidden;
                                DevicesCharm.Visibility = Visibility.Hidden;
                                if (vn3.Content != "0" && vn3.Content != "-1")
                                {
                                    SettingsCharmInactive.Visibility = Visibility.Hidden;
                                }
                                else
                                {
                                    SettingsCharmInactive.Visibility = Visibility.Visible;
                                }

                                if (vn4.Content != "0" && vn4.Content != "-1")
                                {
                                    DevicesCharmInactive.Visibility = Visibility.Hidden;
                                }
                                else
                                {
                                    DevicesCharmInactive.Visibility = Visibility.Visible;
                                }
                                ShareCharmInactive.Visibility = Visibility.Visible;
                                NoColor.Visibility = Visibility.Visible;
                                DevicesCharmInactive.Visibility = Visibility.Visible;
                                SettingsCharmInactive.Visibility = Visibility.Visible;
                            }

                            if (activeIcon == 2)
                            {
                                byte winKey = (byte)KeyInterop.VirtualKeyFromKey(Key.LWin);
                                const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
                                const uint KEYEVENTF_KEYUP = 0x0002;
                                _ = keybd_event(winKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
                                _ = keybd_event(winKey, 0, KEYEVENTF_KEYUP, 0);

                                swipeIn = false;
                                keyboardShortcut = false;
                                charmsAppear = false;
                                charmsUse = false;
                                charmsActivate = false;
                                pokeCharms = false;
                                if (useAnimations == false)
                                {
                                    this.Opacity = 0.000;
                                    CharmsClock.Opacity = 0.000;

                                    var brush = (Brush)converter.ConvertFromString("#00111111");
                                    Background = brush;
                                }

                                mouseIn = false;

                                SearchDown.Visibility = Visibility.Hidden;
                                ShareDown.Visibility = Visibility.Hidden;
                                WinDown.Visibility = Visibility.Hidden;
                                DevicesDown.Visibility = Visibility.Hidden;
                                SettingsDown.Visibility = Visibility.Hidden;

                                SearchText.Visibility = Visibility.Hidden;
                                ShareText.Visibility = Visibility.Hidden;
                                WinText.Visibility = Visibility.Hidden;
                                DevicesText.Visibility = Visibility.Hidden;
                                SettingsText.Visibility = Visibility.Hidden;

                                SearchCharm.Visibility = Visibility.Hidden;
                                ShareCharm.Visibility = Visibility.Hidden;
                                MetroColor.Visibility = Visibility.Hidden;
                                DevicesCharm.Visibility = Visibility.Hidden;
                                SettingsCharm.Visibility = Visibility.Hidden;

                                SearchCharmInactive.Visibility = Visibility.Visible;
                                ShareCharmInactive.Visibility = Visibility.Visible;
                                NoColor.Visibility = Visibility.Visible;
                                if (vn4.Content != "0" && vn4.Content != "-1")
                                {
                                    DevicesCharmInactive.Visibility = Visibility.Hidden;
                                }
                                else
                                {
                                    DevicesCharmInactive.Visibility = Visibility.Visible;
                                }
                                if (vn3.Content != "0" && vn3.Content != "-1")
                                {
                                    SettingsCharmInactive.Visibility = Visibility.Hidden;
                                }
                                else
                                {
                                    SettingsCharmInactive.Visibility = Visibility.Visible;
                                }
                            }

                            if (activeIcon == 3)
                            {
                                findDevices = true;
                                swipeIn = false;
                                keyboardShortcut = false;
                                charmsAppear = false;
                                charmsUse = false;
                                charmsActivate = false;
                                pokeCharms = false;
                                if (useAnimations == false)
                                {
                                    this.Opacity = 0.000;
                                    CharmsClock.Opacity = 0.000;

                                    var brush = (Brush)converter.ConvertFromString("#00111111");
                                    Background = brush;
                                }

                                mouseIn = false;

                                SearchDown.Visibility = Visibility.Hidden;
                                ShareDown.Visibility = Visibility.Hidden;
                                WinDown.Visibility = Visibility.Hidden;
                                DevicesDown.Visibility = Visibility.Hidden;
                                SettingsDown.Visibility = Visibility.Hidden;

                                SearchText.Visibility = Visibility.Hidden;
                                ShareText.Visibility = Visibility.Hidden;
                                WinText.Visibility = Visibility.Hidden;
                                DevicesText.Visibility = Visibility.Hidden;
                                SettingsText.Visibility = Visibility.Hidden;

                                SearchCharm.Visibility = Visibility.Hidden;
                                ShareCharm.Visibility = Visibility.Hidden;
                                MetroColor.Visibility = Visibility.Hidden;
                                DevicesCharm.Visibility = Visibility.Hidden;
                                if (vn3.Content != "0" && vn3.Content != "-1")
                                {
                                    SettingsCharmInactive.Visibility = Visibility.Hidden;
                                }
                                else
                                {
                                    SettingsCharmInactive.Visibility = Visibility.Visible;
                                }

                                if (vn4.Content != "0" && vn4.Content != "-1")
                                {
                                    DevicesCharmInactive.Visibility = Visibility.Hidden;
                                }
                                else
                                {
                                    DevicesCharmInactive.Visibility = Visibility.Visible;
                                }
                                ShareCharmInactive.Visibility = Visibility.Visible;
                                NoColor.Visibility = Visibility.Visible;
                                DevicesCharmInactive.Visibility = Visibility.Visible;
                                SettingsCharmInactive.Visibility = Visibility.Visible;
                            }

                            if (activeIcon == 4)
                            {
                                openSettings = true;
                                swipeIn = false;
                                keyboardShortcut = false;
                                charmsAppear = false;
                                charmsUse = false;
                                charmsActivate = false;
                                pokeCharms = false;
                                if (useAnimations == false)
                                {
                                    this.Opacity = 0.000;
                                    CharmsClock.Opacity = 0.000;

                                    var brush = (Brush)converter.ConvertFromString("#00111111");
                                    Background = brush;
                                }

                                mouseIn = false;

                                SearchDown.Visibility = Visibility.Hidden;
                                ShareDown.Visibility = Visibility.Hidden;
                                WinDown.Visibility = Visibility.Hidden;
                                DevicesDown.Visibility = Visibility.Hidden;
                                SettingsDown.Visibility = Visibility.Hidden;

                                SearchText.Visibility = Visibility.Hidden;
                                ShareText.Visibility = Visibility.Hidden;
                                WinText.Visibility = Visibility.Hidden;
                                DevicesText.Visibility = Visibility.Hidden;
                                SettingsText.Visibility = Visibility.Hidden;

                                SearchCharm.Visibility = Visibility.Hidden;
                                ShareCharm.Visibility = Visibility.Hidden;
                                MetroColor.Visibility = Visibility.Hidden;
                                DevicesCharm.Visibility = Visibility.Hidden;
                                SettingsCharm.Visibility = Visibility.Hidden;

                                SearchCharmInactive.Visibility = Visibility.Visible;
                                ShareCharmInactive.Visibility = Visibility.Visible;
                                NoColor.Visibility = Visibility.Visible;
                                if (vn4.Content != "0" && vn4.Content != "-1")
                                {
                                    DevicesCharmInactive.Visibility = Visibility.Hidden;
                                }
                                else
                                {
                                    DevicesCharmInactive.Visibility = Visibility.Visible;
                                }
                                if (vn3.Content != "0" && vn3.Content != "-1")
                                {
                                    SettingsCharmInactive.Visibility = Visibility.Hidden;
                                }
                                else
                                {
                                    SettingsCharmInactive.Visibility = Visibility.Visible;
                                }
                            }
                            waitTimer = 0;
                        }
                    }

                    if (WinDown.Visibility == Visibility.Visible && WinHover.Visibility == Visibility.Visible && SystemParameters.HighContrast == false)
                    {
                        WinCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Windows8Down.png", UriKind.Relative));
                        var brush = (Brush)converter.ConvertFromString("#444444");
                        FadeBlocker.Background = brush;
                    }

                    if (WinHover.Visibility == Visibility.Visible && WinDown.Visibility != Visibility.Visible && SystemParameters.HighContrast == false)
                    {
                        WinCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Windows8Hover.png", UriKind.Relative));
                        var brush = (Brush)converter.ConvertFromString("#333333");
                        FadeBlocker.Background = brush;
                    }

                    if (WinHover.Visibility != Visibility.Visible && WinDown.Visibility != Visibility.Visible && SystemParameters.HighContrast == false)
                    {
                        WinCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Windows8.png", UriKind.Relative));
                        var brush = (Brush)converter.ConvertFromString("#111111");
                        FadeBlocker.Background = brush;
                    }

                    if (charmsUse == false)
                    {
                        SearchHover.Visibility = Visibility.Hidden;
                        ShareHover.Visibility = Visibility.Hidden;

                        WinCharm.Visibility = Visibility.Visible;
                        WinHover.Visibility = Visibility.Hidden;

                        DevicesHover.Visibility = Visibility.Hidden;
                        SettingsHover.Visibility = Visibility.Hidden;
                    }

                    if (mouseIn == true && activeIcon != 6)
                    {
                        SearchHover.Visibility = Visibility.Hidden;
                        ShareHover.Visibility = Visibility.Hidden;
                        WinCharmUse = false;
                        WinCharm.Visibility = Visibility.Visible;
                        WinHover.Visibility = Visibility.Hidden;
                        DevicesHover.Visibility = Visibility.Hidden;
                        SettingsHover.Visibility = Visibility.Hidden;
                        activeIcon = 6;
                    }
                }
                if (keyboardShortcut == false)
                {
                    if (activeScreen == 0 && usingTouch == false)
                    {
                        if (numVal > mainwidth - 12 & numVal2 < 12 && vn.Content != "1" || numVal > mainwidth - 12 & numVal2 > mainheight - 40 && vn2.Content != "1")
                        {
                            swipeIn = true;
                            charmsTimer += 1;
                            charmsWait += 1;
                        }
                    }

                    if (pokeCharms == true && vn.Content != "1" || pokeCharms == true && vn2.Content != "1")
                    {
                        swipeIn = true;
                        charmsTimer += 1;
                        charmsWait += 1;
                    }

                    if (numVal < mainwidth - 116 && swipeIn == true && useAnimations == true && activeScreen == 0 && charmsUse == true || pokeCharms == true && outofTime == true || forceClose == true)
                    {
                        swipeIn = false;
                    }

                    if (numVal < mainwidth - 116 && useAnimations == false && activeScreen == 0)
                    {
                        charmsTimer = 0;
                        charmsAppear = false;
                        charmsUse = false;
                        charmsActivate = false;
                        pokeCharms = false;
                        mouseIn = false;

                        this.Opacity = 0.000;
                        CharmsClock.Opacity = 0.000;

                        var brush = (Brush)converter.ConvertFromString("#00111111");
                        Background = brush;

                        CharmsClock.Hide();

                        SearchDown.Visibility = Visibility.Hidden;
                        ShareDown.Visibility = Visibility.Hidden;
                        WinDown.Visibility = Visibility.Hidden;
                        DevicesDown.Visibility = Visibility.Hidden;
                        SettingsDown.Visibility = Visibility.Hidden;

                        SearchText.Visibility = Visibility.Hidden;
                        ShareText.Visibility = Visibility.Hidden;
                        WinText.Visibility = Visibility.Hidden;
                        DevicesText.Visibility = Visibility.Hidden;
                        SettingsText.Visibility = Visibility.Hidden;

                        SearchCharm.Visibility = Visibility.Hidden;
                        ShareCharm.Visibility = Visibility.Hidden;
                        MetroColor.Visibility = Visibility.Hidden;
                        DevicesCharm.Visibility = Visibility.Hidden;
                        SettingsCharm.Visibility = Visibility.Hidden;

                        SearchCharmInactive.Visibility = Visibility.Visible;
                        ShareCharmInactive.Visibility = Visibility.Visible;
                        NoColor.Visibility = Visibility.Visible;
                        if (vn4.Content != "0" && vn4.Content != "-1")
                        {
                            DevicesCharmInactive.Visibility = Visibility.Hidden;
                        }
                        else
                        {
                            DevicesCharmInactive.Visibility = Visibility.Visible;
                        }
                        if (vn3.Content != "0" && vn3.Content != "-1")
                        {
                            SettingsCharmInactive.Visibility = Visibility.Hidden;
                        }
                        else
                        {
                            SettingsCharmInactive.Visibility = Visibility.Visible;
                        }
                    }

                    if (charmsAppear == true)
                    {
                        charmsTimer = 1945;
                    }

                    if (activeScreen > 0 && usingTouch == false)
                    {
                        if (numVal > screenwidth - 12 & numVal2 < 12 && vn.Content != "1" || numVal > screenwidth - 12 & numVal2 > screenheight - 40 && vn2.Content != "1")
                        {
                            swipeIn = true;
                            charmsTimer += 1;
                            charmsWait += 1;
                        }
                    }

                    if (pokeCharms == true && vn.Content != "1" || pokeCharms == true && vn2.Content != "1")
                    {
                        swipeIn = true;
                        charmsTimer += 1;
                        charmsWait += 1;
                    }

                    if (numVal < screenwidth - 116 && swipeIn == true && useAnimations == true && activeScreen > 0 || pokeCharms == true && outofTime == true || forceClose == true)
                    {
                        swipeIn = false;
                    }

                    if (numVal < screenwidth - 116 && useAnimations == false && activeScreen > 0)
                    {
                        charmsTimer = 0;
                        charmsAppear = false;
                        charmsUse = false;
                        charmsActivate = false;
                        pokeCharms = false;
                        mouseIn = false;

                        this.Opacity = 0.000;
                        CharmsClock.Opacity = 0.000;

                        var brush = (Brush)converter.ConvertFromString("#00111111");
                        Background = brush;

                        CharmsClock.Hide();

                        SearchDown.Visibility = Visibility.Hidden;
                        ShareDown.Visibility = Visibility.Hidden;
                        WinDown.Visibility = Visibility.Hidden;
                        DevicesDown.Visibility = Visibility.Hidden;
                        SettingsDown.Visibility = Visibility.Hidden;

                        SearchText.Visibility = Visibility.Hidden;
                        ShareText.Visibility = Visibility.Hidden;
                        WinText.Visibility = Visibility.Hidden;
                        DevicesText.Visibility = Visibility.Hidden;
                        SettingsText.Visibility = Visibility.Hidden;

                        SearchCharm.Visibility = Visibility.Hidden;
                        ShareCharm.Visibility = Visibility.Hidden;
                        MetroColor.Visibility = Visibility.Hidden;
                        DevicesCharm.Visibility = Visibility.Hidden;
                        SettingsCharm.Visibility = Visibility.Hidden;

                        SearchCharmInactive.Visibility = Visibility.Visible;
                        ShareCharmInactive.Visibility = Visibility.Visible;
                        NoColor.Visibility = Visibility.Visible;
                        if (vn4.Content != "0" && vn4.Content != "-1")
                        {
                            DevicesCharmInactive.Visibility = Visibility.Hidden;
                        }
                        else
                        {
                            DevicesCharmInactive.Visibility = Visibility.Visible;
                        }
                        if (vn3.Content != "0" && vn3.Content != "-1")
                        {
                            SettingsCharmInactive.Visibility = Visibility.Hidden;
                        }
                        else
                        {
                            SettingsCharmInactive.Visibility = Visibility.Visible;
                        }
                    }

                    if (charmsAppear == true)
                    {
                        charmsTimer = 1945;
                    }
                }

                if (numVal < mainwidth - 116 && swipeIn == true && keyboardShortcut == false & activeScreen == 0)
                {
                    charmsAppear = false;
                    charmsUse = false;
                    charmsActivate = false;
                    pokeCharms = false;
                    if (useAnimations == false)
                    {
                        this.Opacity = 0.000;
                        CharmsClock.Opacity = 0.000;

                        var brush = (Brush)converter.ConvertFromString("#00111111");
                        Background = brush;

                        CharmsClock.Hide();
                    }
                    mouseIn = false;

                    SearchDown.Visibility = Visibility.Hidden;
                    ShareDown.Visibility = Visibility.Hidden;
                    WinDown.Visibility = Visibility.Hidden;
                    DevicesDown.Visibility = Visibility.Hidden;
                    SettingsDown.Visibility = Visibility.Hidden;

                    SearchText.Visibility = Visibility.Hidden;
                    ShareText.Visibility = Visibility.Hidden;
                    WinText.Visibility = Visibility.Hidden;
                    DevicesText.Visibility = Visibility.Hidden;
                    SettingsText.Visibility = Visibility.Hidden;

                    SearchCharm.Visibility = Visibility.Hidden;
                    ShareCharm.Visibility = Visibility.Hidden;
                    MetroColor.Visibility = Visibility.Hidden;
                    DevicesCharm.Visibility = Visibility.Hidden;
                    SettingsCharm.Visibility = Visibility.Hidden;

                    SearchCharmInactive.Visibility = Visibility.Visible;
                    ShareCharmInactive.Visibility = Visibility.Visible;
                    NoColor.Visibility = Visibility.Visible;
                    if (vn4.Content != "0" && vn4.Content != "-1")
                    {
                        DevicesCharmInactive.Visibility = Visibility.Hidden;
                    }
                    else
                    {
                        DevicesCharmInactive.Visibility = Visibility.Visible;
                    }
                    if (vn3.Content != "0" && vn3.Content != "-1")
                    {
                        SettingsCharmInactive.Visibility = Visibility.Hidden;
                    }
                    else
                    {
                        SettingsCharmInactive.Visibility = Visibility.Visible;
                    }
                }

                if (activeScreen == 0 && usingTouch == false)
                {
                    if (numVal > mainwidth - 15 && numVal2 < 12 && vn.Content != "1" && forceClose == false || numVal > mainwidth - 15 & numVal2 > mainheight - 40 && vn2.Content != "1" && forceClose == false)
                    {
                        blockRepeating = 0;
                        preventReload = true;
                        pokeCharms = true;
                        holder = true;
                    }
                    else
                    {
                        preventReload = false;
                        holder = false;
                    }

                    if (numVal < mainwidth - 116 || forceClose == true)
                    {
                        blockRepeating = 0;
                        pokeCharms = false;
                        preventReload = false;
                    }
                }

                if (activeScreen > 0 && usingTouch == false)
                {
                    if (numVal > screenwidth - 15 && numVal2 < 12 && vn.Content != "1" && forceClose == false || numVal > screenwidth - 15 & numVal2 > screenheight - 40 && vn2.Content != "1" && forceClose == false)
                    {
                        blockRepeating = 0;
                        preventReload = true;
                        pokeCharms = true;
                        holder = true;
                    }
                    else
                    {
                        preventReload = false;
                        holder = false;
                    }

                    if (numVal < screenwidth - 116 || forceClose == true)
                    {
                        blockRepeating = 0;
                        pokeCharms = false;
                        preventReload = false;
                    }
                }

                if (System.Windows.Forms.Control.MouseButtons == MouseButtons.None && pokeCharms == true && charmsTimer > charmsDelay && keyboardShortcut == false && forceClose == false || System.Windows.Forms.Control.MouseButtons == MouseButtons.None && charmsAppear == true && keyboardShortcut == false && forceClose == false || keyboardShortcut == true && forceClose == false)
                {
                    if (useAnimations == true)
                    {

                        if (textSearch != 0)
                        {
                            textSearch -= 10;
                        }

                        if (textShare != 0)
                        {
                            textShare -= 10;
                        }

                        if (textWin != 0)
                        {
                            textWin -= 10;
                        }

                        if (textDevices != 0)
                        {
                            textDevices -= 10;
                        }

                        if (textSettings != 0)
                        {
                            textSettings -= 10;
                        }

                        if (scrollSearch != 0)
                        {
                            scrollSearch -= 10;
                        }

                        if (scrollShare != 0)
                        {
                            scrollShare -= 10;
                        }

                        if (scrollWin != 0)
                        {
                            scrollWin -= 10;
                        }

                        if (scrollDevices != 0)
                        {
                            scrollDevices -= 10;
                        }

                        if (scrollSettings != 0)
                        {
                            scrollSettings -= 10;
                        }

                        if (keyboardShortcut == true && scrollWin >= -10 && winStretch >= 80.31)
                        {
                            winStretch -= 1.00;
                        }
                    }
                    charmsAppear = true;

                    if (charmsAppear == true && charmsUse == true)
                    {
                        SearchCharmInactive.Visibility = Visibility.Hidden;
                        ShareCharmInactive.Visibility = Visibility.Hidden;
                        NoColor.Visibility = Visibility.Hidden;
                        WinCharmInactive.Visibility = Visibility.Hidden;
                        DevicesCharmInactive.Visibility = Visibility.Hidden;
                        SettingsCharmInactive.Visibility = Visibility.Hidden;
                    }
                }

                //FIRING UP !!
                if (charmsUse == false && usingTouch == false)
                {
                    if (activeScreen == 0)
                    {
                        if (System.Windows.Forms.Control.MouseButtons == MouseButtons.None && pokeCharms == true & numVal2 > 208 & numVal2 < mainheight - 702 && keyboardShortcut == false && swipeIn == true && useAnimations == true && outofTime == false && forceClose == false || System.Windows.Forms.Control.MouseButtons == MouseButtons.None && pokeCharms == true & numVal2 > 193 & numVal2 < mainheight - 202 && keyboardShortcut == false && swipeIn == true && useAnimations == true && outofTime == false && forceClose == false || System.Windows.Forms.Control.MouseButtons == MouseButtons.None && pokeCharms == true & numVal2 > 208 & numVal2 < mainheight - 702 && keyboardShortcut == false && useAnimations == false && outofTime == false && forceClose == false || System.Windows.Forms.Control.MouseButtons == MouseButtons.None && pokeCharms == true & numVal2 > 193 & numVal2 < mainheight - 202 && keyboardShortcut == false && useAnimations == false && outofTime == false && forceClose == false)
                        {
                            charmsActivate = true;
                        }
                        else
                        {
                            charmsActivate = false;
                        }
                    }

                    if (activeScreen > 0)
                    {
                        if (System.Windows.Forms.Control.MouseButtons == MouseButtons.None && pokeCharms == true & numVal2 > 208 & numVal2 < screenheight - 702 && keyboardShortcut == false && swipeIn == true && useAnimations == true && outofTime == false && forceClose == false || System.Windows.Forms.Control.MouseButtons == MouseButtons.None && pokeCharms == true & numVal2 > 193 & numVal2 < screenheight - 202 && keyboardShortcut == false && swipeIn == true && useAnimations == true && outofTime == false && forceClose == false || System.Windows.Forms.Control.MouseButtons == MouseButtons.None && pokeCharms == true & numVal2 > 208 & numVal2 < screenheight - 702 && keyboardShortcut == false && useAnimations == false && outofTime == false && forceClose == false || System.Windows.Forms.Control.MouseButtons == MouseButtons.None && pokeCharms == true & numVal2 > 193 & numVal2 < screenheight - 202 && keyboardShortcut == false && useAnimations == false && outofTime == false && forceClose == false)
                        {
                            pokeCharms = true;
                            charmsAppear = true;
                            charmsActivate = true;
                            charmsUse = true;
                            keyboardShortcut = true;
                            this.BringIntoView();
                            this.Focus();
                            this.Activate();
                        }
                        else
                        {
                            charmsActivate = false;
                        }
                    }
                }

                if (usingTouch == true)
                {
                    if (System.Windows.Forms.Control.MouseButtons == MouseButtons.None & numVal > mainwidth - 75 && numVal < mainwidth)
                    {
                        scrollSearch = 0;
                        scrollShare = 0;
                        scrollWin = 0;
                        scrollDevices = 0;
                        scrollSettings = 0;

                        textSearch = 0;
                        textShare = 0;
                        textWin = 0;
                        textDevices = 0;
                        textSettings = 0;
                        charmsAppear = true;
                        charmsTimer = 1;
                        charmsWait = 1;
                    }
                    SearchCharmInactive.Visibility = Visibility.Hidden;
                    ShareCharmInactive.Visibility = Visibility.Hidden;
                    WinCharmInactive.Visibility = Visibility.Hidden;
                    NoColor.Visibility = Visibility.Hidden;
                    DevicesCharmInactive.Visibility = Visibility.Hidden;
                    SettingsCharmInactive.Visibility = Visibility.Hidden;

                    if (activeScreen == 0)
                    {

                        if (System.Windows.Forms.Control.MouseButtons == MouseButtons.Left & numVal > mainwidth - 75 && numVal < mainwidth - 10 && pokeCharms == false && this.IsActive == false)
                        {
                            pokeCharms = true;
                            charmsActivate = true;
                            this.Opacity = 1;
                            SearchText.Visibility = Visibility.Visible;
                            ShareText.Visibility = Visibility.Visible;
                            WinText.Visibility = Visibility.Visible;
                            DevicesText.Visibility = Visibility.Visible;
                            SettingsText.Visibility = Visibility.Visible;

                            SearchCharm.Visibility = Visibility.Visible;
                            ShareCharm.Visibility = Visibility.Visible;
                            MetroColor.Visibility = Visibility.Visible;
                            WinCharm.Visibility = Visibility.Visible;
                            if (vn4.Content != "0" && vn4.Content != "-1")
                            {
                                DevicesCharm.Visibility = Visibility.Hidden;
                            }
                            else
                            {
                                DevicesCharm.Visibility = Visibility.Visible;
                            }
                            if (vn3.Content != "0" && vn3.Content != "-1")
                            {
                                SettingsCharm.Visibility = Visibility.Hidden;
                            }
                            else
                            {
                                SettingsCharm.Visibility = Visibility.Visible;
                            }

                            this.BringIntoView();
                            this.Focus();
                            this.Activate();
                        }
                    }
                    //sanic
                    if (activeScreen > 0)
                    {
                        if (System.Windows.Forms.Control.MouseButtons == MouseButtons.Left & numVal > screenwidth - 75 && numVal < screenwidth - 10 && keyboardShortcut == false && pokeCharms == false && this.IsActive == false)
                        {
                            pokeCharms = true;
                            charmsActivate = true;
                            this.Opacity = 1;
                            SearchText.Visibility = Visibility.Visible;
                            ShareText.Visibility = Visibility.Visible;
                            WinText.Visibility = Visibility.Visible;
                            DevicesText.Visibility = Visibility.Visible;
                            SettingsText.Visibility = Visibility.Visible;

                            SearchCharm.Visibility = Visibility.Visible;
                            ShareCharm.Visibility = Visibility.Visible;
                            MetroColor.Visibility = Visibility.Visible;
                            WinCharm.Visibility = Visibility.Visible;
                            if (vn4.Content != "0" && vn4.Content != "-1")
                            {
                                DevicesCharm.Visibility = Visibility.Hidden;
                            }
                            else
                            {
                                DevicesCharm.Visibility = Visibility.Visible;
                            }
                            if (vn3.Content != "0" && vn3.Content != "-1")
                            {
                                SettingsCharm.Visibility = Visibility.Hidden;
                            }
                            else
                            {
                                SettingsCharm.Visibility = Visibility.Visible;
                            }

                            this.BringIntoView();
                            this.Focus();
                            this.Activate();
                        }
                    }
                }

                //FIX!!!
                if (charmsAppear == true && charmsUse == true)
                {
                    if (activeScreen != 0)
                    {
                        if (SystemParameters.HighContrast == false)
                        {
                            var brush = (Brush)converter.ConvertFromString("#ff111111");
                            Background = brush;
                        }
                        else
                        {
                            Background = SystemColors.WindowBrush;
                        }
                        this.Opacity = IHOb;
                    }

                    mouseIn = false;
                    SearchText.Visibility = Visibility.Visible;
                    ShareText.Visibility = Visibility.Visible;
                    WinText.Visibility = Visibility.Visible;
                    DevicesText.Visibility = Visibility.Visible;
                    SettingsText.Visibility = Visibility.Visible;

                    SearchCharm.Visibility = Visibility.Visible;
                    ShareCharm.Visibility = Visibility.Visible;
                    MetroColor.Visibility = Visibility.Visible;
                    WinCharm.Visibility = Visibility.Visible;
                    if (vn4.Content != "0" && vn4.Content != "-1")
                    {
                        DevicesCharm.Visibility = Visibility.Hidden;
                    }
                    else
                    {
                        DevicesCharm.Visibility = Visibility.Visible;
                    }
                    if (vn3.Content != "0" && vn3.Content != "-1")
                    {
                        SettingsCharm.Visibility = Visibility.Hidden;
                    }
                    else
                    {
                        SettingsCharm.Visibility = Visibility.Visible;
                    }

                    SearchCharmInactive.Visibility = Visibility.Hidden;
                    ShareCharmInactive.Visibility = Visibility.Hidden;
                    WinCharmInactive.Visibility = Visibility.Hidden;
                    NoColor.Visibility = Visibility.Hidden;
                    DevicesCharmInactive.Visibility = Visibility.Hidden;
                    SettingsCharmInactive.Visibility = Visibility.Hidden;
                    charmsActivate = true;
                    charmsUse = true;
                }

                //activate without the animations...
                if (charmsActivate == true && useAnimations == false || Keyboard.IsKeyDown(Key.LWin) && Keyboard.IsKeyDown(Key.C) && useAnimations == false || Keyboard.IsKeyDown(Key.RWin) && Keyboard.IsKeyDown(Key.C) && useAnimations == false)
                {
                    this.Focus();
                    this.Activate();
                    this.BringIntoView();
                    if (noClocks.Content == "-1" || noClocks.Content == "0")
                    {
                        CharmsClock.Show();
                    }

                    if (charmsAppear == false)
                    {
                        charmsAppear = true;
                    }

                    if (SystemParameters.HighContrast == false)
                    {
                        var brush = (Brush)converter.ConvertFromString("#ff111111");
                        Background = brush;
                    }
                    else
                    {
                        Background = SystemColors.WindowBrush;
                    }

                    mouseIn = false;
                    SearchText.Visibility = Visibility.Visible;
                    ShareText.Visibility = Visibility.Visible;
                    WinText.Visibility = Visibility.Visible;
                    DevicesText.Visibility = Visibility.Visible;
                    SettingsText.Visibility = Visibility.Visible;

                    SearchCharm.Visibility = Visibility.Visible;
                    ShareCharm.Visibility = Visibility.Visible;
                    MetroColor.Visibility = Visibility.Visible;
                    WinCharm.Visibility = Visibility.Visible;
                    if (vn4.Content != "0" && vn4.Content != "-1")
                    {
                        DevicesCharm.Visibility = Visibility.Hidden;
                    }
                    else
                    {
                        DevicesCharm.Visibility = Visibility.Visible;
                    }
                    if (vn3.Content != "0" && vn3.Content != "-1")
                    {
                        SettingsCharm.Visibility = Visibility.Hidden;
                    }
                    else
                    {
                        SettingsCharm.Visibility = Visibility.Visible;
                    }

                    SearchCharmInactive.Visibility = Visibility.Hidden;
                    ShareCharmInactive.Visibility = Visibility.Hidden;
                    WinCharmInactive.Visibility = Visibility.Hidden;
                    NoColor.Visibility = Visibility.Hidden;
                    DevicesCharmInactive.Visibility = Visibility.Hidden;
                    SettingsCharmInactive.Visibility = Visibility.Hidden;

                    if (Keyboard.IsKeyDown(Key.LWin) && Keyboard.IsKeyDown(Key.C) || Keyboard.IsKeyDown(Key.RWin) && Keyboard.IsKeyDown(Key.C))
                    {

                        FadeBlocker.Opacity = 1.0;
                        if (activeIcon == 6)
                        {
                            activeIcon = 2;
                        }

                        //Search highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 0)
                        {
                            SearchHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 0)
                        {
                            SearchHover.Visibility = Visibility.Hidden;
                        }

                        //Share highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 1)
                        {
                            ShareHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 1)
                        {
                            ShareHover.Visibility = Visibility.Hidden;
                        }

                        //Start highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 2)
                        {
                            WinCharmUse = true;
                            WinHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 2)
                        {
                            WinCharmUse = false;
                            WinHover.Visibility = Visibility.Hidden;
                        }

                        //Devices highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 3)
                        {
                            DevicesHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 3)
                        {
                            DevicesHover.Visibility = Visibility.Hidden;
                        }

                        //Settings highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 4)
                        {
                            SettingsHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 4)
                        {
                            SettingsHover.Visibility = Visibility.Hidden;
                        }

                    }

                    charmsUse = true;
                }

                //activate with the animations!!
                if (charmsActivate == true && useAnimations == true || Keyboard.IsKeyDown(Key.LWin) && Keyboard.IsKeyDown(Key.C) && useAnimations == true || Keyboard.IsKeyDown(Key.RWin) && Keyboard.IsKeyDown(Key.C) && useAnimations == true)
                {
                    this.Focus();
                    this.Activate();
                    this.BringIntoView();

                    if (charmsAppear == false)
                    {
                        charmsAppear = true;
                    }

                    if (keyboardShortcut == true)
                    {
                        MetroColor.Opacity = 1.0;
                        WinCharm.Opacity = 1.0;
                        CharmBG.Opacity = 0.000;
                    }

                    if (scrollSearch != 0 && scrollSearch < 199 && useAnimations == true)
                    {
                        if (SystemParameters.HighContrast == false)
                        {
                            var brush = (Brush)converter.ConvertFromString("#ff111111");
                            Background = brush;
                        }
                        else
                        {

                            Background = SystemColors.WindowBrush;
                        }
                    }

                    mouseIn = false;
                    SearchText.Visibility = Visibility.Visible;
                    ShareText.Visibility = Visibility.Visible;
                    WinText.Visibility = Visibility.Visible;
                    DevicesText.Visibility = Visibility.Visible;
                    SettingsText.Visibility = Visibility.Visible;

                    SearchCharm.Visibility = Visibility.Visible;
                    ShareCharm.Visibility = Visibility.Visible;
                    MetroColor.Visibility = Visibility.Visible;
                    WinCharm.Visibility = Visibility.Visible;
                    if (vn4.Content != "0" && vn4.Content != "-1")
                    {
                        DevicesCharm.Visibility = Visibility.Hidden;
                    }
                    else
                    {
                        DevicesCharm.Visibility = Visibility.Visible;
                    }
                    if (vn3.Content != "0" && vn3.Content != "-1")
                    {
                        SettingsCharm.Visibility = Visibility.Hidden;
                    }
                    else
                    {
                        SettingsCharm.Visibility = Visibility.Visible;
                    }

                    SearchCharmInactive.Visibility = Visibility.Hidden;
                    ShareCharmInactive.Visibility = Visibility.Hidden;
                    WinCharmInactive.Visibility = Visibility.Hidden;
                    NoColor.Visibility = Visibility.Hidden;
                    DevicesCharmInactive.Visibility = Visibility.Hidden;
                    SettingsCharmInactive.Visibility = Visibility.Hidden;

                    if (Keyboard.IsKeyDown(Key.LWin) && Keyboard.IsKeyDown(Key.C) || Keyboard.IsKeyDown(Key.RWin) && Keyboard.IsKeyDown(Key.C))
                    {
                        FadeBlocker.Opacity = 1.0;
                        if (activeIcon == 6)
                        {
                            activeIcon = 2;
                        }

                        //Search highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 0)
                        {
                            SearchHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 0)
                        {
                            SearchHover.Visibility = Visibility.Hidden;
                        }

                        //Share highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 1)
                        {
                            ShareHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 1)
                        {
                            ShareHover.Visibility = Visibility.Hidden;
                        }

                        //Start highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 2)
                        {
                            WinCharmUse = true;
                            WinHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 2)
                        {
                            WinCharmUse = false;
                            WinHover.Visibility = Visibility.Hidden;
                        }

                        //Devices highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 3)
                        {
                            DevicesHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 3)
                        {
                            DevicesHover.Visibility = Visibility.Hidden;
                        }

                        //Settings highlighted
                        if (mouseIn == false && keyboardShortcut == true && activeIcon == 4)
                        {
                            SettingsHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 4)
                        {
                            SettingsHover.Visibility = Visibility.Hidden;
                        }

                        this.Opacity = IHOb;

                        if (noClocks.Content == "-1" || noClocks.Content == "0")
                        {
                            CharmsClock.Show();
                        }

                        CharmsClock.Opacity = IHOb;

                        keyboardShortcut = true;
                    }
                    else
                    {
                        if (useAnimations == false)
                        {
                            this.Opacity = 1.0;

                            if (noClocks.Content == "-1" || noClocks.Content == "0")
                            {
                                CharmsClock.Show();
                            }
                            CharmsClock.Opacity = 1.0;
                        }

                        if (useAnimations == true)
                        {
                            this.Opacity = IHOb;

                            if (noClocks.Content == "-1" || noClocks.Content == "0")
                            {
                                CharmsClock.Show();
                            }

                            CharmsClock.Opacity = IHOb;
                        }
                    }

                    charmsUse = true;
                }

                if (Keyboard.IsKeyDown(Key.Escape) && charmsUse == true)
                {
                    if (useAnimations == false)
                    {
                        swipeIn = false;
                        keyboardShortcut = false;
                        charmsAppear = false;
                        charmsUse = false;
                        charmsActivate = false;
                        pokeCharms = false;

                        this.Opacity = 0.000;
                        CharmsClock.Opacity = 0.000;

                        var brush = (Brush)converter.ConvertFromString("#00111111");
                        Background = brush;
                        CharmsClock.Hide();

                        mouseIn = false;

                        SearchDown.Visibility = Visibility.Hidden;
                        ShareDown.Visibility = Visibility.Hidden;
                        WinDown.Visibility = Visibility.Hidden;
                        DevicesDown.Visibility = Visibility.Hidden;
                        SettingsDown.Visibility = Visibility.Hidden;

                        SearchText.Visibility = Visibility.Hidden;
                        ShareText.Visibility = Visibility.Hidden;
                        WinText.Visibility = Visibility.Hidden;
                        DevicesText.Visibility = Visibility.Hidden;
                        SettingsText.Visibility = Visibility.Hidden;

                        SearchCharm.Visibility = Visibility.Hidden;
                        ShareCharm.Visibility = Visibility.Hidden;
                        MetroColor.Visibility = Visibility.Hidden;
                        DevicesCharm.Visibility = Visibility.Hidden;
                        SettingsCharm.Visibility = Visibility.Hidden;

                        SearchCharmInactive.Visibility = Visibility.Visible;
                        ShareCharmInactive.Visibility = Visibility.Visible;
                        NoColor.Visibility = Visibility.Visible;
                        if (vn4.Content != "0" && vn4.Content != "-1")
                        {
                            DevicesCharmInactive.Visibility = Visibility.Hidden;
                        }
                        else
                        {
                            DevicesCharmInactive.Visibility = Visibility.Visible;
                        }
                        if (vn3.Content != "0" && vn3.Content != "-1")
                        {
                            SettingsCharmInactive.Visibility = Visibility.Hidden;
                        }
                        else
                        {
                            SettingsCharmInactive.Visibility = Visibility.Visible;
                        }
                    }
                    else
                    {
                        escKey = true;
                    }
                }
                //High Contrast support!
                if (SystemParameters.HighContrast == false)
                {
                    WinFader.Visibility = Visibility.Visible;
                    FadeBlocker.Visibility = Visibility.Visible;
                    CharmBorder.Visibility = Visibility.Hidden;
                    CharmBorder.Background = (Brush)converter.ConvertFromString("#111111");
                    SearchText.Foreground = (Brush)converter.ConvertFromString("#a0a0a0");
                    ShareText.Foreground = (Brush)converter.ConvertFromString("#a0a0a0");
                    WinText.Foreground = (Brush)converter.ConvertFromString("#a0a0a0");
                    DevicesText.Foreground = (Brush)converter.ConvertFromString("#a0a0a0");
                    SettingsText.Foreground = (Brush)converter.ConvertFromString("#a0a0a0");
                    CharmBG.Background = (Brush)converter.ConvertFromString("#111111");
                    if (MetroColor.Background.ToString() == "#00000000")
                    {
                        SearchCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Search.png", UriKind.Relative));
                        ShareCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Share.png", UriKind.Relative));
                        WinCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Windows8.png", UriKind.Relative));
                        DevicesCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Devices.png", UriKind.Relative));
                        SettingsCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Settings.png", UriKind.Relative));
                        MetroColor.Background = SystemParameters.WindowGlassBrush;
                    }
                    SearchHover.Background = (Brush)converter.ConvertFromString("#333333");
                    ShareHover.Background = (Brush)converter.ConvertFromString("#333333");
                    WinHover.Background = (Brush)converter.ConvertFromString("#333333");
                    DevicesHover.Background = (Brush)converter.ConvertFromString("#333333");
                    SettingsHover.Background = (Brush)converter.ConvertFromString("#333333");
                    SearchDown.Background = (Brush)converter.ConvertFromString("#444444");
                    ShareDown.Background = (Brush)converter.ConvertFromString("#444444");
                    WinDown.Background = (Brush)converter.ConvertFromString("#444444");
                    DevicesDown.Background = (Brush)converter.ConvertFromString("#444444");
                    SettingsDown.Background = (Brush)converter.ConvertFromString("#444444");
                }

                if (SystemParameters.HighContrast == true)
                {
                    WinFader.Visibility = Visibility.Hidden;
                    FadeBlocker.Visibility = Visibility.Hidden;
                    CharmBorder.Visibility = Visibility.Visible;
                    CharmBorder.Background = SystemColors.WindowTextBrush;
                    SearchText.Foreground = SystemColors.WindowTextBrush;
                    ShareText.Foreground = SystemColors.WindowTextBrush;
                    WinText.Foreground = SystemColors.WindowTextBrush;
                    DevicesText.Foreground = SystemColors.WindowTextBrush;
                    SettingsText.Foreground = SystemColors.WindowTextBrush;
                    CharmBG.Background = SystemColors.WindowBrush;
                    System.Drawing.Color col = System.Drawing.ColorTranslator.FromHtml(SystemColors.WindowBrush.ToString());
                    if (col.R * 0.2126 + col.G * 0.7152 + col.B * 0.0722 < 255 / 2)
                    {
                        // dark color
                        SearchCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Search.png", UriKind.Relative));
                        ShareCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Share.png", UriKind.Relative));
                        WinCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Windows8HighContrast.png", UriKind.Relative));
                        DevicesCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Devices.png", UriKind.Relative));
                        SettingsCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Settings.png", UriKind.Relative));
                    }
                    else
                    {
                        // light color
                        SearchCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/SearchDark.png", UriKind.Relative));
                        ShareCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/ShareDark.png", UriKind.Relative));
                        WinCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Windows8HighContrastDark.png", UriKind.Relative));
                        DevicesCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/DevicesDark.png", UriKind.Relative));
                        SettingsCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/SettingsDark.png", UriKind.Relative));
                    }

                    MetroColor.Background = (Brush)converter.ConvertFromString("#00000000");
                    SearchHover.Background = SystemColors.HighlightBrush;
                    ShareHover.Background = SystemColors.HighlightBrush;
                    WinHover.Background = SystemColors.HighlightBrush;
                    DevicesHover.Background = SystemColors.HighlightBrush;
                    SettingsHover.Background = SystemColors.HighlightBrush;
                    SearchDown.Background = SystemColors.WindowBrush;
                    ShareDown.Background = SystemColors.WindowBrush;
                    WinDown.Background = SystemColors.WindowBrush;
                    DevicesDown.Background = SystemColors.WindowBrush;
                    SettingsDown.Background = SystemColors.WindowBrush;
                    if (keyboardShortcut == false)
                    {
                        CharmBorder.Opacity = CharmBG.Opacity;
                    }
                    else
                    {
                        CharmBorder.Opacity = 1;
                    }
                }

                if (SearchHover.Visibility == Visibility.Visible && SystemParameters.HighContrast == true)
                {
                    SearchText.Foreground = SystemColors.HighlightTextBrush;
                    System.Drawing.Color col = System.Drawing.ColorTranslator.FromHtml(SystemColors.HighlightBrush.ToString());
                    if (col.R * 0.2126 + col.G * 0.7152 + col.B * 0.0722 < 255 / 2)
                    {
                        // dark color
                        SearchCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Search.png", UriKind.Relative));
                    }
                    else
                    {
                        // light color
                        SearchCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/SearchDark.png", UriKind.Relative));
                    }
                }

                if (ShareHover.Visibility == Visibility.Visible && SystemParameters.HighContrast == true)
                {
                    ShareText.Foreground = SystemColors.HighlightTextBrush;
                    System.Drawing.Color col = System.Drawing.ColorTranslator.FromHtml(SystemColors.HighlightBrush.ToString());
                    if (col.R * 0.2126 + col.G * 0.7152 + col.B * 0.0722 < 255 / 2)
                    {
                        // dark color
                        ShareCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Share.png", UriKind.Relative));
                    }
                    else
                    {
                        // light color
                        ShareCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/ShareDark.png", UriKind.Relative));
                    }
                }

                if (WinHover.Visibility == Visibility.Visible && SystemParameters.HighContrast == true)
                {
                    WinText.Foreground = SystemColors.HighlightTextBrush;
                    System.Drawing.Color col = System.Drawing.ColorTranslator.FromHtml(SystemColors.HighlightBrush.ToString());
                    if (col.R * 0.2126 + col.G * 0.7152 + col.B * 0.0722 < 255 / 2)
                    {
                        // dark color
                        WinCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Windows8HighContrast.png", UriKind.Relative));
                    }
                    else
                    {
                        // light color
                        WinCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Windows8HighContrastDark.png", UriKind.Relative));
                    }
                }

                if (DevicesHover.Visibility == Visibility.Visible && SystemParameters.HighContrast == true)
                {
                    DevicesText.Foreground = SystemColors.HighlightTextBrush;
                    System.Drawing.Color col = System.Drawing.ColorTranslator.FromHtml(SystemColors.HighlightBrush.ToString());
                    if (col.R * 0.2126 + col.G * 0.7152 + col.B * 0.0722 < 255 / 2)
                    {
                        // dark color
                        DevicesCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Devices.png", UriKind.Relative));
                    }
                    else
                    {
                        // light color
                        DevicesCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/DevicesDark.png", UriKind.Relative));
                    }
                }

                if (SettingsHover.Visibility == Visibility.Visible && SystemParameters.HighContrast == true)
                {
                    SettingsText.Foreground = SystemColors.HighlightTextBrush;
                    System.Drawing.Color col = System.Drawing.ColorTranslator.FromHtml(SystemColors.HighlightBrush.ToString());
                    if (col.R * 0.2126 + col.G * 0.7152 + col.B * 0.0722 < 255 / 2)
                    {
                        // dark color
                        SettingsCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Settings.png", UriKind.Relative));
                    }
                    else
                    {
                        // light color
                        SettingsCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/SettingsDark.png", UriKind.Relative));
                    }
                }

                if (SearchDown.Visibility == Visibility.Visible && SystemParameters.HighContrast == true)
                {
                    SearchText.Foreground = SystemColors.InactiveCaptionTextBrush;
                }

                if (ShareDown.Visibility == Visibility.Visible && SystemParameters.HighContrast == true)
                {
                    ShareText.Foreground = SystemColors.InactiveCaptionTextBrush;
                }

                if (WinDown.Visibility == Visibility.Visible && SystemParameters.HighContrast == true)
                {
                    WinText.Foreground = SystemColors.InactiveCaptionTextBrush;
                }

                if (DevicesDown.Visibility == Visibility.Visible && SystemParameters.HighContrast == true)
                {
                    DevicesText.Foreground = SystemColors.InactiveCaptionTextBrush;
                }

                if (SettingsDown.Visibility == Visibility.Visible && SystemParameters.HighContrast == true)
                {
                    SettingsText.Foreground = SystemColors.InactiveCaptionTextBrush;
                }

                if (SearchHover.Visibility == Visibility.Hidden && SearchDown.Visibility == Visibility.Hidden && SystemParameters.HighContrast == true)
                {
                    SearchText.Foreground = SystemColors.WindowTextBrush;
                }

                if (ShareHover.Visibility == Visibility.Hidden && ShareDown.Visibility == Visibility.Hidden && SystemParameters.HighContrast == true)
                {
                    ShareText.Foreground = SystemColors.WindowTextBrush;
                }

                if (WinHover.Visibility == Visibility.Hidden && WinDown.Visibility == Visibility.Hidden && SystemParameters.HighContrast == true)
                {
                    WinText.Foreground = SystemColors.WindowTextBrush;
                }

                if (DevicesHover.Visibility == Visibility.Hidden && DevicesDown.Visibility == Visibility.Hidden && SystemParameters.HighContrast == true)
                {
                    DevicesText.Foreground = SystemColors.WindowTextBrush;
                }

                if (SettingsHover.Visibility == Visibility.Hidden && SettingsDown.Visibility == Visibility.Hidden && SystemParameters.HighContrast == true)
                {
                    SettingsText.Foreground = SystemColors.WindowTextBrush;
                }

//fixes for the Windows shell

                if (this.IsActive == false && charmsUse == true && isGui == true)
                    {
                                byte escKey = (byte)KeyInterop.VirtualKeyFromKey(Key.Escape);
                                const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
                                const uint KEYEVENTF_KEYUP = 0x0002;
                                _ = keybd_event(escKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
                                _ = keybd_event(escKey, 0, KEYEVENTF_KEYUP, 0);

        IntPtr lHwnd = FindWindow("Shell_TrayWnd", null);
        SendMessage(lHwnd, WM_COMMAND, (IntPtr)MIN_ALL, IntPtr.Zero);  
        System.Threading.Thread.Sleep(700); //Hopefully, this will fix the bug.
        SendMessage(lHwnd, WM_COMMAND, (IntPtr)MIN_ALL_UNDO, IntPtr.Zero); 

                    SetForegroundWindow(Int32.Parse(ActiveCharm.Content.ToString()));
                    this.Focus();
                    this.Activate();
                    this.BringIntoView();
                    charmsUse = true;
                    isGui = false;
                }

                //metro app interrogation.
                if (isMetro == true && this.IsActive == false && CharmsClock.IsVisible == false)
                {
                    IntPtr mWnd = GetForegroundWindow();

                    RECT rct = new RECT();
                    GetWindowRect(mWnd, ref rct);

                    rctTop = rct.Top;
                    rctLeft = rct.Left;
                }

                if (isMetro == false && this.IsActive == false && CharmsClock.IsVisible == false)
                {
                    IntPtr mWnd = GetForegroundWindow();

                    rctTop = 0;
                    rctLeft = 0;
                }

                //End of Charms Bar Code
            }));
        }
        private void Charms_MouseMove(object sender, System.EventArgs e)
        {
            cursorStay = 0;
            mouseIn = true;
            twoInputs = true;

            //active!
            if (searchActive == true && searchHover == true)
            {
                SearchDown.Visibility = Visibility.Visible;
                SearchHover.Visibility = Visibility.Hidden;
            }

            if (shareActive == true && shareHover == true)
            {
                ShareDown.Visibility = Visibility.Visible;
                ShareHover.Visibility = Visibility.Hidden;
            }

            if (winActive == true && winHover == true && SystemParameters.HighContrast == false)
            {
                WinCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Windows8Down.png", UriKind.Relative));
                var brush = (Brush)converter.ConvertFromString("#444444");
                FadeBlocker.Background = brush;
                WinDown.Visibility = Visibility.Visible;
                WinHover.Visibility = Visibility.Hidden;
            }

            if (winActive == true && winHover == true && SystemParameters.HighContrast == true)
            {
                WinDown.Visibility = Visibility.Visible;
                WinHover.Visibility = Visibility.Hidden;
            }

            if (devicesActive == true && devicesHover == true)
            {
                DevicesDown.Visibility = Visibility.Visible;
                DevicesHover.Visibility = Visibility.Hidden;
            }

            if (settingsActive == true && settingsHover == true)
            {
                SettingsDown.Visibility = Visibility.Visible;
                SettingsHover.Visibility = Visibility.Hidden;
            }

            //not active...
            if (searchActive == true && searchHover == false)
            {
                SearchDown.Visibility = Visibility.Hidden;
                SearchHover.Visibility = Visibility.Hidden;
            }

            if (shareActive == true && shareHover == false)
            {
                ShareDown.Visibility = Visibility.Hidden;
                ShareHover.Visibility = Visibility.Hidden;
            }

            if (winActive == true && winHover == false && SystemParameters.HighContrast == false)
            {
                WinCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Windows8.png", UriKind.Relative));
                var brush = (Brush)converter.ConvertFromString("#111111");
                FadeBlocker.Background = brush;
                WinDown.Visibility = Visibility.Hidden;
                WinHover.Visibility = Visibility.Hidden;
            }

            if (devicesActive == true && devicesHover == false)
            {
                DevicesDown.Visibility = Visibility.Hidden;
                DevicesHover.Visibility = Visibility.Hidden;
            }

            if (settingsActive == true && settingsHover == false)
            {
                SettingsDown.Visibility = Visibility.Hidden;
                SettingsHover.Visibility = Visibility.Hidden;
            }

            if (this.IsActive == false && charmsUse == false && keyboardShortcut == false)
            {
                if (useAnimations == false)
                {
                    this.Opacity = 1.0;
                    CharmsClock.Opacity = 0.000;

                    var brush = (Brush)converter.ConvertFromString("#00111111");
                    Background = brush;
                }
                mouseIn = false;
            }

            if (this.IsActive == true && charmsUse == false && keyboardShortcut == false)
            {
                if (useAnimations == false)
                {
                    this.Opacity = 1.0;
                    CharmsClock.Opacity = 0.000;
                }
            }

            if (this.IsActive == true && charmsUse == true && keyboardShortcut == false)
            {
                this.Focus();
                this.Activate();
                this.BringIntoView();
                if (useAnimations == false)
                {
                    this.Opacity = 1.0;
                    CharmsClock.Opacity = 1.0;
                }
            }
        }

        private void Charms_MouseUp(object sender, System.EventArgs e)
        {
            SearchDown.Visibility = Visibility.Hidden;
            ShareDown.Visibility = Visibility.Hidden;
            WinDown.Visibility = Visibility.Hidden;
            DevicesDown.Visibility = Visibility.Hidden;
            SettingsDown.Visibility = Visibility.Hidden;
        }

        private void Charms_MouseDown(object sender, System.EventArgs e)
        {
            mouseIn = true;
            twoInputs = true;

            if (this.IsActive == true && charmsUse == false && keyboardShortcut == false)
            {
                if (useAnimations == false)
                {
                    this.Opacity = 0.000;
                    CharmsClock.Opacity = 0.000;
                }
            }

            if (this.IsActive == true && charmsUse == true && keyboardShortcut == false)
            {
                if (useAnimations == false)
                {
                    this.Opacity = 1.0;
                    CharmsClock.Opacity = 1.0;
                }
            }
        }

        private void Search_MouseDown(object sender, System.EventArgs e)
        {
            searchActive = true;
            SearchDown.Visibility = Visibility.Visible;
        }

        private void Share_MouseDown(object sender, System.EventArgs e)
        {
            shareActive = true;
            ShareDown.Visibility = Visibility.Visible;
        }

        private void Win_MouseDown(object sender, System.EventArgs e)
        {
            winActive = true;
            if (SystemParameters.HighContrast == false)
            {
                WinCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Windows8Down.png", UriKind.Relative));
                var brush = (Brush)converter.ConvertFromString("#444444");
                FadeBlocker.Background = brush;
                WinDown.Visibility = Visibility.Visible;
            }

            if (SystemParameters.HighContrast == true)
            {
                WinDown.Visibility = Visibility.Visible;
            }
        }

        private void Devices_MouseDown(object sender, System.EventArgs e)
        {
            devicesActive = true;
            DevicesDown.Visibility = Visibility.Visible;
        }

        private void Settings_MouseDown(object sender, System.EventArgs e)
        {
            settingsActive = true;
            SettingsDown.Visibility = Visibility.Visible;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;

            public static implicit operator Point(POINT point)
            {
                return new Point(point.X, point.Y);
            }
        }

        public static Point ElementPointToScreenPoint(UIElement element, Point pointOnElement)
        {
            return element.PointToScreen(pointOnElement);
        }

        private void Search_MouseEnter(object sender, System.EventArgs e)
        {
            searchHover = true;
            mouseIn = true;
            twoInputs = true;
            if (charmsAppear == true && System.Windows.Forms.Control.MouseButtons != MouseButtons.Left)
            {
                SearchHover.Visibility = Visibility.Visible;
            }
        }

        private void Search_MouseLeave(object sender, System.EventArgs e)
        {
            searchHover = false;
            mouseIn = true;
            twoInputs = true;
            if (charmsAppear == true && System.Windows.Forms.Control.MouseButtons != MouseButtons.Left)
            {
                SearchHover.Visibility = Visibility.Hidden;
            }

            if (charmsAppear == true && System.Windows.Forms.Control.MouseButtons == MouseButtons.Left)
            {
                SearchDown.Visibility = Visibility.Visible;
            }
        }

        private void Share_MouseEnter(object sender, System.EventArgs e)
        {
            shareHover = true;
            mouseIn = true;
            twoInputs = true;
            if (charmsAppear == true && System.Windows.Forms.Control.MouseButtons != MouseButtons.Left)
            {
                ShareHover.Visibility = Visibility.Visible;
            }

            if (charmsAppear == true && System.Windows.Forms.Control.MouseButtons == MouseButtons.Left)
            {
                ShareDown.Visibility = Visibility.Visible;
            }
        }

        private void Share_MouseLeave(object sender, System.EventArgs e)
        {
            shareHover = false;
            mouseIn = true;
            twoInputs = true;
            if (charmsAppear == true && System.Windows.Forms.Control.MouseButtons != MouseButtons.Left)
            {
                ShareHover.Visibility = Visibility.Hidden;
            }
        }

        private void Win_MouseEnter(object sender, System.EventArgs e)
        {
            winHover = true;
            mouseIn = true;
            twoInputs = true;
            WinCharmUse = true;
            if (charmsAppear == true && System.Windows.Forms.Control.MouseButtons != MouseButtons.Left && SystemParameters.HighContrast == false)
            {
                WinCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Windows8Hover.png", UriKind.Relative));
                var brush = (Brush)converter.ConvertFromString("#333333");
                FadeBlocker.Background = brush;
                WinHover.Visibility = Visibility.Visible;
            }

            if (charmsAppear == true && System.Windows.Forms.Control.MouseButtons == MouseButtons.Left && SystemParameters.HighContrast == false)
            {
                WinCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Windows8Down.png", UriKind.Relative));
                var brush = (Brush)converter.ConvertFromString("#444444");
                FadeBlocker.Background = brush;
                WinDown.Visibility = Visibility.Visible;
            }

            if (charmsAppear == true && System.Windows.Forms.Control.MouseButtons != MouseButtons.Left && SystemParameters.HighContrast == true)
            {
                WinHover.Visibility = Visibility.Visible;
            }

            if (WinCharm.IsMouseOver == true && charmsAppear == true && System.Windows.Forms.Control.MouseButtons == MouseButtons.Left && SystemParameters.HighContrast == true)
            {
                WinDown.Visibility = Visibility.Visible;
            }
        }

        private void Win_MouseLeave(object sender, System.EventArgs e)
        {
            winHover = false;
            mouseIn = true;
            twoInputs = true;
            WinCharmUse = false;
            if (charmsAppear == true)
            {
                if (System.Windows.Forms.Control.MouseButtons == MouseButtons.None && SystemParameters.HighContrast == false)
                {
                    WinCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Windows8.png", UriKind.Relative));
                    var brush = (Brush)converter.ConvertFromString("#111111");
                    FadeBlocker.Background = brush;
                    WinDown.Visibility = Visibility.Hidden;
                }

                if (System.Windows.Forms.Control.MouseButtons == MouseButtons.Left && SystemParameters.HighContrast == false)
                {
                    WinCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Windows8Down.png", UriKind.Relative));
                    var brush = (Brush)converter.ConvertFromString("#444444");
                    FadeBlocker.Background = brush;
                }

                if (CharmsHover.Content == "False" && System.Windows.Forms.Control.MouseButtons == MouseButtons.Left && SystemParameters.HighContrast == false)
                {
                    WinCharm.Source = new BitmapImage(new Uri(@"/Assets/Images/Windows8.png", UriKind.Relative));
                    var brush = (Brush)converter.ConvertFromString("#111111");
                    FadeBlocker.Background = brush;
                    WinDown.Visibility = Visibility.Hidden;
                }

                if (charmsAppear == true && System.Windows.Forms.Control.MouseButtons != MouseButtons.Left)
                {
                    WinHover.Visibility = Visibility.Hidden;
                }
                WinDown.Visibility = Visibility.Hidden;
            }
        }

        private void Devices_MouseEnter(object sender, System.EventArgs e)
        {
            devicesHover = true;
            mouseIn = true;
            twoInputs = true;
            if (charmsAppear == true && System.Windows.Forms.Control.MouseButtons != MouseButtons.Left)
            {
                DevicesHover.Visibility = Visibility.Visible;
            }

            if (charmsAppear == true && System.Windows.Forms.Control.MouseButtons == MouseButtons.Left)
            {
                DevicesDown.Visibility = Visibility.Visible;
            }
        }

        private void Devices_MouseLeave(object sender, System.EventArgs e)
        {
            devicesHover = false;
            mouseIn = true;
            twoInputs = true;
            if (charmsAppear == true && System.Windows.Forms.Control.MouseButtons != MouseButtons.Left)
            {
                DevicesHover.Visibility = Visibility.Hidden;
            }
            DevicesDown.Visibility = Visibility.Hidden;
        }

        private void Settings_MouseEnter(object sender, System.EventArgs e)
        {
            settingsHover = true;
            mouseIn = true;
            twoInputs = true;
            if (charmsAppear == true && System.Windows.Forms.Control.MouseButtons != MouseButtons.Left)
            {
                SettingsHover.Visibility = Visibility.Visible;
            }

            if (charmsAppear == true && System.Windows.Forms.Control.MouseButtons == MouseButtons.Left)
            {
                SettingsDown.Visibility = Visibility.Visible;
            }
        }

        private void Settings_MouseLeave(object sender, System.EventArgs e)
        {
            settingsHover = false;
            mouseIn = true;
            twoInputs = true;
            if (charmsAppear == true && System.Windows.Forms.Control.MouseButtons != MouseButtons.Left)
            {
                SettingsHover.Visibility = Visibility.Hidden;
            }
            SettingsDown.Visibility = Visibility.Hidden;
        }
        private void Charms_MouseEnter(object sender, System.EventArgs e)
        {
            mouseIn = true;
            twoInputs = true;
            if (this.IsActive == false && charmsUse == true && keyboardShortcut == false)
            {
                SearchDown.Visibility = Visibility.Hidden;
                ShareDown.Visibility = Visibility.Hidden;
                WinDown.Visibility = Visibility.Hidden;
                DevicesDown.Visibility = Visibility.Hidden;
                SettingsDown.Visibility = Visibility.Hidden;

                SearchText.Visibility = Visibility.Hidden;
                ShareText.Visibility = Visibility.Hidden;
                WinText.Visibility = Visibility.Hidden;
                DevicesText.Visibility = Visibility.Hidden;
                SettingsText.Visibility = Visibility.Hidden;

                SearchCharm.Visibility = Visibility.Hidden;
                ShareCharm.Visibility = Visibility.Hidden;
                MetroColor.Visibility = Visibility.Hidden;
                DevicesCharm.Visibility = Visibility.Hidden;
                SettingsCharm.Visibility = Visibility.Hidden;

                SearchCharmInactive.Visibility = Visibility.Visible;
                ShareCharmInactive.Visibility = Visibility.Visible;
                NoColor.Visibility = Visibility.Visible;
                if (vn4.Content != "0" && vn4.Content != "-1")
                {
                    DevicesCharmInactive.Visibility = Visibility.Hidden;
                }
                else
                {
                    DevicesCharmInactive.Visibility = Visibility.Visible;
                }
                if (vn3.Content != "0" && vn3.Content != "-1")
                {
                    SettingsCharmInactive.Visibility = Visibility.Hidden;
                }
                else
                {
                    SettingsCharmInactive.Visibility = Visibility.Visible;
                }

                charmsUse = false;
            }

            if (this.IsActive == true && charmsUse == false && keyboardShortcut == false)
            {
                if (useAnimations == false)
                {
                    this.Opacity = 1.0;
                    CharmsClock.Opacity = 1.0;
                }
            }

            if (this.IsActive == true && charmsUse == true && keyboardShortcut == false)
            {
                if (useAnimations == false)
                {
                    this.Opacity = 1.0;
                    CharmsClock.Opacity = 1.0;
                }
            }
        }

        private void Charms_MouseLeave(object sender, System.EventArgs e)
        {
            mouseIn = true;
            SearchDown.Visibility = Visibility.Hidden;
            ShareDown.Visibility = Visibility.Hidden;
            WinDown.Visibility = Visibility.Hidden;
            DevicesDown.Visibility = Visibility.Hidden;
            SettingsDown.Visibility = Visibility.Hidden;

            searchHover = false;
            shareHover = false;
            winHover = false;
            devicesHover = false;
            settingsHover = false;

            searchActive = false;
            shareActive = false;
            winActive = false;
            devicesActive = false;
            settingsActive = false;

            if (this.IsActive == false && charmsUse == true && twoInputs == false)
            {
                if (useAnimations == false)
                {
                    var brush = (Brush)converter.ConvertFromString("#00111111");
                    Background = brush;
                }

                mouseIn = false;
                charmsUse = false;
            }

            if (this.IsActive == true && charmsUse == false && twoInputs == false)
            {
                if (useAnimations == false)
                {
                    this.Opacity = 0.000;
                    CharmsClock.Opacity = 0.000;
                }
            }

            if (this.IsActive == false && charmsUse == false && twoInputs == false)
            {
                if (useAnimations == false)
                {
                    this.Opacity = 0.000;
                    CharmsClock.Opacity = 0.000;
                }
            }

            if (this.IsActive == false && charmsUse == true && twoInputs == false)
            {
                charmsUse = false;
                if (useAnimations == false)
                {
                    this.Opacity = 0.000;
                    CharmsClock.Opacity = 0.000;
                }
            }

            if (this.IsActive == false && charmsUse == true && twoInputs == true)
            {
                charmsUse = false;
                if (useAnimations == false)
                {
                    this.Opacity = 0.000;
                    CharmsClock.Opacity = 0.000;
                }
            }

            charmsTimer = 0;
        }

        // Handle the UI exceptions by showing a dialog box, and asking the user whether
        // or not they wish to abort execution.
        private static void Form1_UIThreadException(object sender, ThreadExceptionEventArgs t)
        {
            System.Windows.Forms.Application.Restart();
        }
    }
}