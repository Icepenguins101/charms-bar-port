using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Timers;
using System.Threading;
using System.Windows.Threading;
using System.Windows.Media.Media3D;
using System.Reflection.Metadata;
using System.Reflection;
using System.Xml.Linq;
using System.ComponentModel;
using Microsoft.Win32;

namespace CharmsBarPort
{
    public partial class CharmsBar : Window
    {
        BrushConverter converter = new();
        Window CharmsClock = new CharmsClock();

        public int restoreWindow = 0;
        public bool swipeIn = false;
        public bool keyboardShortcut = false;
        public bool charmsAppear = false;
        public bool charmsUse = false;
        public int charmsTimer = 0;
        public int charmsWait = 0;
        public bool WinCharmUse = false;
        public int activeIcon = 2;
        public bool mouseIn = false;
        public bool twoInputs = false;
        public int waitTimer = 0;
        public int keyboardTimer = 0;
        public bool useAnimations = false;
        public double IHOb = 1.0;
        public bool escKey = false;
        public int dasSlide = 0;

        //Supports Windows 8.1 registry hacks!
        public bool noTopRight = false;
        public bool noBottomRight = false;

        //For the animations!
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

        public CharmsBar()
        {
            // Here we obtain the current primary screen resolution. 
            var dispWidth = SystemParameters.PrimaryScreenWidth;
            var dispHeight = SystemParameters.PrimaryScreenHeight;
            Topmost = true;
            ShowInTaskbar = false;
            WindowStyle = WindowStyle.None;
            ResizeMode = ResizeMode.NoResize;
            AllowsTransparency = true;
            Height = dispHeight;
            Width = 86; // 80 to 90 is magic spot.
            WindowStartupLocation = WindowStartupLocation.Manual;
            Left = dispWidth - Width;
            Top = 0;
            //BrushConverter converter = new();
            var brush = (Brush)converter.ConvertFromString("#00111111");
            Background = brush;
            Opacity = 0.001;
            SystemParameters.StaticPropertyChanged += this.SystemParameters_StaticPropertyChanged;
            this.Loaded += ControlLoaded;
            CharmsClock.Hide();
            InitializeComponent();
            this.Closing += new System.ComponentModel.CancelEventHandler(MainWindow_Closing);

        }

        void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
        }

        protected override void OnActivated(EventArgs e)
        {
            if (this.IsActive == true && charmsUse == true)
            {
                if (useAnimations == false)
                {
                    this.Opacity = 1.0;
                    CharmsClock.Opacity = 1.0;
                    CharmsClock.Show();
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
                if (useAnimations == false)
                {
                    this.Opacity = 0.001;
                    CharmsClock.Opacity = 0.001;

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

                if (charmsTimer == 0 && keyboardShortcut == false)
                {
                    swipeIn = false;
                    twoInputs = false;
                    keyboardShortcut = false;
                    charmsAppear = false;
                    if (useAnimations == false)
                    {
                        this.Opacity = 0.001;
                        CharmsClock.Opacity = 0.001;

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

        private void Search_MouseUp(object sender, MouseButtonEventArgs e)
        {
            byte winKey = (byte)KeyInterop.VirtualKeyFromKey(Key.LWin);
            byte sKey = (byte)KeyInterop.VirtualKeyFromKey(Key.S);
            const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
            const uint KEYEVENTF_KEYUP = 0x0002;
            charmsUse = false;
            _ = keybd_event(winKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
            _ = keybd_event(sKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
            _ = keybd_event(winKey, 0, KEYEVENTF_KEYUP, 0);
            _ = keybd_event(sKey, 0, KEYEVENTF_KEYUP, 0);

            swipeIn = false;
            keyboardShortcut = false;
            charmsAppear = false;
            charmsUse = false;
            if (useAnimations == false)
            {
                this.Opacity = 0.001;
                CharmsClock.Opacity = 0.001;

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
        private void Share_MouseUp(object sender, MouseButtonEventArgs e)
        {
            byte printScreenKey = (byte)KeyInterop.VirtualKeyFromKey(Key.PrintScreen);
            const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
            const uint KEYEVENTF_KEYUP = 0x0002;
            charmsUse = false;
            _ = keybd_event(printScreenKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
            _ = keybd_event(printScreenKey, 0, KEYEVENTF_KEYUP, 0);

            swipeIn = false;
            keyboardShortcut = false;
            charmsAppear = false;
            charmsUse = false;
            if (useAnimations == false)
            {
                this.Opacity = 0.001;
                CharmsClock.Opacity = 0.001;

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

        /// <summary>
        /// Start button click logic
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Win_MouseUp(object sender, MouseButtonEventArgs e)
        {
            byte winKey = (byte)KeyInterop.VirtualKeyFromKey(Key.LWin);
            const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
            const uint KEYEVENTF_KEYUP = 0x0002;
            charmsUse = false;
            _ = keybd_event(winKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
            _ = keybd_event(winKey, 0, KEYEVENTF_KEYUP, 0);

            swipeIn = false;
            keyboardShortcut = false;
            charmsAppear = false;
            charmsUse = false;
            if (useAnimations == false)
            {
                this.Opacity = 0.001;
                CharmsClock.Opacity = 0.001;
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
            byte winKey = (byte)KeyInterop.VirtualKeyFromKey(Key.LWin);
            byte pKey = (byte)KeyInterop.VirtualKeyFromKey(Key.P);
            const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
            const uint KEYEVENTF_KEYUP = 0x0002;
            charmsUse = false;
            _ = keybd_event(winKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
            _ = keybd_event(pKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
            _ = keybd_event(winKey, 0, KEYEVENTF_KEYUP, 0);
            _ = keybd_event(pKey, 0, KEYEVENTF_KEYUP, 0);

            swipeIn = false;
            keyboardShortcut = false;
            charmsAppear = false;
            charmsUse = false;
            if (useAnimations == false)
            {
                this.Opacity = 0.001;
                CharmsClock.Opacity = 0.001;

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

        private void Settings_MouseUp(object sender, MouseButtonEventArgs e)
        {
            byte winKey = (byte)KeyInterop.VirtualKeyFromKey(Key.LWin);
            byte iKey = (byte)KeyInterop.VirtualKeyFromKey(Key.I);
            const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
            const uint KEYEVENTF_KEYUP = 0x0002;
            charmsUse = false;
            _ = keybd_event(winKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
            _ = keybd_event(iKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
            _ = keybd_event(winKey, 0, KEYEVENTF_KEYUP, 0);
            _ = keybd_event(iKey, 0, KEYEVENTF_KEYUP, 0);

            swipeIn = false;
            keyboardShortcut = false;
            charmsAppear = false;
            charmsUse = false;
            if (useAnimations == false)
            {
                this.Opacity = 0.001;
                CharmsClock.Opacity = 0.001;

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

        private System.Timers.Timer t = null;
        private readonly Dispatcher dispatcher = Dispatcher.CurrentDispatcher;
        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

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

        private void _initTimer()
        {
            t = new System.Timers.Timer();
            t.Interval = 10;
            t.Elapsed += OnTimedEvent;
            t.AutoReset = true;
            t.Enabled = true;
            t.Start();
        }

        private void OnTimedEvent(object sender, ElapsedEventArgs e)
        {
            dispatcher.BeginInvoke((Action)(() =>
            {
                    Mouse.Capture(this);
                    Point pointToWindow = Mouse.GetPosition(this);
                    Point pointToScreen = PointToScreen(pointToWindow);
                    Mouse.Capture(null);

                    int numVal = Int32.Parse(pointToScreen.X.ToString());
                    int numVal2 = Int32.Parse(pointToScreen.Y.ToString());

                if (useAnimations == false)
                {
                    CharmBG.Opacity = 0.001;
                    CharmBG.Visibility = Visibility.Hidden;
                }
                else
                {
                    CharmBG.Visibility = Visibility.Visible;
                }

                if (charmsUse == false)
                {
                    if (useAnimations == true)
                    {
                        if (CharmBG.Opacity > 0.001 && outofTime == false)
                        {
                            FadeBlocker.Opacity -= 0.1;
                            CharmBG.Opacity -= 0.1;
                            WinCharm.Opacity -= 0.1;
                            MetroColor.Opacity -= 0.1;
                        }
                        dasSlide = -300;
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

                WinFader.Margin = new Thickness(dasSlide, 0, 0, 0);
                try
                {
                    RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\ImmersiveShell\\EdgeUi", false);
                    if (key != null)
                    {
                        // get value
                        string vn = key.GetValue("DisableTRCorner", -1, RegistryValueOptions.None).ToString();
                        string vn2 = key.GetValue("DisableBRCorner", -1, RegistryValueOptions.None).ToString();

                        if (vn == "0" || vn == null)
                        {
                            noTopRight = false;
                        }

                        if (vn == "1")
                        {
                            noTopRight = true;
                        }

                        if (vn2 == "0" || vn2 == null)
                        {
                            noBottomRight = false;
                        }

                        if (vn2 == "1")
                        {
                            noBottomRight = true;
                        }
                        // close key
                        key.Close();
                    }
                }
                catch (Exception ex)  //just for demonstration...it's always best to handle specific exceptions
                {
                    //react appropriately
                }

                var dispWidth = SystemParameters.PrimaryScreenWidth;
                var dispHeight = SystemParameters.PrimaryScreenHeight;

                if (IHOb <= 0.001 && useAnimations == true && charmsTimer == 1 || keyboardShortcut == true && IHOb <= 0.001 && useAnimations == true)
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

                if (IHOb <= 0.001 && useAnimations == false && charmsTimer == 1 || keyboardShortcut == true && IHOb <= 0.001 && useAnimations == false)
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
                        useAnimations = false;
                    }
                }
                catch (Win32Exception ex)
                {
                    //error
                }
//animations :)
                if (useAnimations == true)
                {
                    if (charmsUse == false && IHOb < 0.012)
                    {
                        CharmsClock.Hide();
                    }

                    if (charmsUse == true)
                    {
                        CharmsClock.Show();
                    }

                    if (WinCharm.Visibility == Visibility.Hidden && charmsUse == false)
                    {
                        var brush = (Brush)converter.ConvertFromString("#00111111");
                        Background = brush;
                    }

                    if (charmsAppear == true && numVal < dispWidth - 86 && keyboardShortcut == false || charmsAppear == true && keyboardShortcut == true && this.IsActive == false || charmsAppear == true && keyboardShortcut == true && escKey == true || this.IsActive == true && AnyKeyPressed() == true && swipeIn == true && numVal > dispWidth - 85 || ignoreMouseIn == true && numVal > dispWidth - 85 || outofTime == true && numVal > dispWidth - 85)
                    {
                        ignoreMouseIn = true;
                        IHOb -= 0.141;
                        SearchCharmInactive.Opacity = IHOb;
                        ShareCharmInactive.Opacity = IHOb;
                        NoColor.Opacity = IHOb;
                        DevicesCharmInactive.Opacity = IHOb;
                        SettingsCharmInactive.Opacity = IHOb;
                    }

                    if (charmsAppear == true && IHOb < 0.001 && keyboardShortcut == true)
                    {
                        keyboardShortcut = false;
                    }

                    if (charmsAppear == false && IHOb > 0.001 && keyboardShortcut == false || swipeIn == true && numVal < dispWidth - 86 || charmsAppear == true && escKey == true && keyboardShortcut == true || outofTime == true && numVal > dispWidth - 85)
                    {
                        if (escKey == false)
                        {
                            IHOb -= 0.061;
                        }

                        if (escKey == true)
                        {
                            IHOb -= 0.021;
                        }
                        this.Opacity = IHOb;
                        CharmsClock.Opacity = IHOb;
                    }

                    if (charmsAppear == true && numVal < dispWidth - 86 && IHOb < 0.001 && keyboardShortcut == false && outofTime == false && ignoreMouseIn == false)
                    {
                        IHOb = 1.0;
                        SearchCharmInactive.Opacity = IHOb;
                        ShareCharmInactive.Opacity = IHOb;
                        NoColor.Opacity = IHOb;
                        DevicesCharmInactive.Opacity = IHOb;
                        SettingsCharmInactive.Opacity = IHOb;
                        this.Opacity = 0.001;
                        charmsAppear = false;
                        charmsTimer = 0;
                    }

                    if (IHOb < 0.001 && keyboardShortcut == false)
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
                        CharmsClock.Opacity = 0.001;
                        CharmsClock.Hide();
                    }

                    if (charmsAppear == true && numVal > dispWidth - 85 && keyboardShortcut == false && IHOb < 1.0 && charmsUse == false && ignoreMouseIn == false)
                    {
                        IHOb += 0.10;
                        SearchCharmInactive.Opacity = IHOb;
                        ShareCharmInactive.Opacity = IHOb;
                        NoColor.Opacity = IHOb;
                        DevicesCharmInactive.Opacity = IHOb;
                        SettingsCharmInactive.Opacity = IHOb;
                    }

                    if (charmsUse == false && IHOb < 0.001)
                    {
                        IHOb = 0.001;
                        escKey = false;
                    }

                    if (charmsAppear == true && IHOb < 1.1 && ignoreMouseIn == false)
                    {
                        IHOb += 0.05;
                        this.Opacity = IHOb;
                        CharmsClock.Opacity = IHOb;
                    }

                    if (charmsWait > 300 && charmsUse == false)
                    {
                        outofTime = true;
                    }

                    if (charmsWait < 300 && numVal < dispWidth - 86)
                    {
                        outofTime = false;
                    }

                    if (this.Opacity < 0.1)
                    {
                        charmsWait = 0;
                    }

                    if (CharmBG.Opacity < 1.1 && charmsUse == true && charmsAppear == true && useAnimations == true && outofTime == false)
                    {
                        if (keyboardShortcut == false)
                        {
                            WinCharm.Opacity += 0.1;
                            MetroColor.Opacity += 0.1;
                            FadeBlocker.Opacity += 0.1;
                            CharmBG.Opacity += 0.1;
                        }
                    }

                    var searchDas = new Thickness(scrollSearch, -11, 12, -66);
                    SearchCharmInactive.Margin = searchDas;

                    var shareDas = new Thickness(scrollShare-1, 14, 12, -66);
                    ShareCharmInactive.Margin = shareDas;

                    var winDas = new Thickness(scrollWin, -50, 14, 4);
                    NoColor.Margin = winDas;

                    var deviceDas = new Thickness(scrollDevices, -38, 12, -99);
                    DevicesCharmInactive.Margin = deviceDas;

                    var settingsDas = new Thickness(scrollSettings, 14, 12, -99);
                    SettingsCharmInactive.Margin = settingsDas;

                    var searchDas2 = new Thickness(scrollSearch, -11, 12, -66);
                    SearchCharm.Margin = searchDas2;

                    var shareDas2 = new Thickness(scrollShare-1, 14, 12, -66);
                    ShareCharm.Margin = shareDas2;

                    var winDas2 = new Thickness(scrollWin, 11, 12, 4);
                    MetroColor.Margin = winDas2;

                    var deviceDas2 = new Thickness(scrollDevices, 13, 12, -10);
                    DevicesCharm.Margin = deviceDas2;

                    var settingsDas2 = new Thickness(scrollSettings, 14, 12, -99);
                    SettingsCharm.Margin = settingsDas2;

                    var searchDas3 = new Thickness(textSearch+1, 38, 13.141, -44.89);
                    SearchText.Margin = searchDas3;

                    var shareDas3 = new Thickness(textShare, 59, 12, 0);
                    ShareText.Margin = shareDas3;

                    var winDas3 = new Thickness(textWin, -15, 12, 0);
                    WinText.Margin = winDas3;

                    var deviceDas3 = new Thickness(textDevices, 7, 12, 0);
                    DevicesText.Margin = deviceDas3;

                    var settingsDas3 = new Thickness(textSettings, 59, 12, -63);
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
                //no animations :(
                if (useAnimations == false)
                {
                    if (charmsUse == false && IHOb < 0.012)
                    {
                        CharmsClock.Hide();
                    }

                    if (charmsUse == true)
                    {
                        CharmsClock.Show();
                    }

                    if (WinCharm.Visibility == Visibility.Hidden && charmsUse == false)
                    {
                        var brush = (Brush)converter.ConvertFromString("#00111111");
                        Background = brush;
                    }

                    if (charmsAppear == true && numVal < dispWidth - 86 && keyboardShortcut == false || charmsAppear == true && keyboardShortcut == true && this.IsActive == false || charmsAppear == true && keyboardShortcut == true && escKey == true || this.IsActive == true && AnyKeyPressed() == true && swipeIn == true && numVal > dispWidth - 85 || ignoreMouseIn == true && numVal > dispWidth - 85 || outofTime == true && numVal > dispWidth - 85)
                    {
                        ignoreMouseIn = true;
                        IHOb = 0.001;
                        SearchCharmInactive.Opacity = IHOb;
                        ShareCharmInactive.Opacity = IHOb;
                        NoColor.Opacity = IHOb;
                        DevicesCharmInactive.Opacity = IHOb;
                        SettingsCharmInactive.Opacity = IHOb;
                    }

                    if (charmsAppear == true && IHOb < 0.001 && keyboardShortcut == true)
                    {
                        keyboardShortcut = false;
                    }

                    if (charmsAppear == false && IHOb > 0.001 && keyboardShortcut == false || swipeIn == true && numVal < dispWidth - 86 || charmsAppear == true && escKey == true && keyboardShortcut == true || outofTime == true && numVal > dispWidth - 85)
                    {
                        charmsAppear = false;
                        charmsTimer = 0;
                        charmsWait = 0;
                        IHOb = 0.001;
                        this.Opacity = IHOb;
                        CharmsClock.Opacity = IHOb;

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

                    if (charmsAppear == true && numVal < dispWidth - 86 && IHOb < 0.001 && keyboardShortcut == false && outofTime == false)
                    {
                        IHOb = 1.0;
                        SearchCharmInactive.Opacity = IHOb;
                        ShareCharmInactive.Opacity = IHOb;
                        NoColor.Opacity = IHOb;
                        DevicesCharmInactive.Opacity = IHOb;
                        SettingsCharmInactive.Opacity = IHOb;
                        this.Opacity = 0.001;
                        charmsAppear = false;
                        charmsTimer = 0;
                    }

                    if (IHOb == 0.001 && numVal < dispWidth - 86)
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

                        swipeIn = false;
                        charmsUse = false;
                        charmsAppear = false;
                        charmsTimer = 0;
                        charmsWait = 0;
                    }

                    if (IHOb < 0.001 && keyboardShortcut == false)
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
                        keyboardShortcut = false;
                        CharmsClock.Opacity = 0.001;
                        CharmsClock.Hide();
                    }

                    if (charmsAppear == true && numVal > dispWidth - 85 && keyboardShortcut == false && IHOb < 1.0 && charmsUse == false)
                    {
                        IHOb = 1.0;
                        SearchCharmInactive.Opacity = IHOb;
                        ShareCharmInactive.Opacity = IHOb;
                        NoColor.Opacity = IHOb;
                        DevicesCharmInactive.Opacity = IHOb;
                        SettingsCharmInactive.Opacity = IHOb;
                    }

                    if (charmsUse == false && IHOb < 0.001)
                    {
                        IHOb = 0.001;
                        escKey = false;
                    }

                    if (charmsAppear == true && IHOb < 1.1)
                    {
                        IHOb = 1.0;
                        this.Opacity = IHOb;
                        CharmsClock.Opacity = IHOb;
                    }

                    if (charmsWait > 300 && charmsUse == false)
                    {
                        outofTime = true;
                    }

                    if (charmsWait < 300 && numVal < dispWidth - 86)
                    {
                        outofTime = false;
                    }

                    if (CharmBG.Opacity == 0.001 && charmsUse == true && charmsAppear == true && outofTime == false)
                    {
                        if (keyboardShortcut == false)
                        {
                            WinCharm.Opacity = 1.0;
                            MetroColor.Opacity = 1.0;
                            FadeBlocker.Opacity = 1.0;
                            CharmBG.Opacity = 1.0;
                        }
                    }

                    var searchDas = new Thickness(0, -11, 12, -66);
                    SearchCharmInactive.Margin = searchDas;

                    var shareDas = new Thickness(0 - 1, 14, 12, -66);
                    ShareCharmInactive.Margin = shareDas;

                    var winDas = new Thickness(0, -50, 14, 4);
                    NoColor.Margin = winDas;

                    var deviceDas = new Thickness(0, -38, 12, -99);
                    DevicesCharmInactive.Margin = deviceDas;

                    var settingsDas = new Thickness(0, 14, 12, -99);
                    SettingsCharmInactive.Margin = settingsDas;

                    var searchDas2 = new Thickness(0, -11, 12, -66);
                    SearchCharm.Margin = searchDas2;

                    var shareDas2 = new Thickness(0 - 1, 14, 12, -66);
                    ShareCharm.Margin = shareDas2;

                    var winDas2 = new Thickness(0, 11, 12, 4);
                    MetroColor.Margin = winDas2;

                    var deviceDas2 = new Thickness(0, 13, 12, -10);
                    DevicesCharm.Margin = deviceDas2;

                    var settingsDas2 = new Thickness(0, 14, 12, -99);
                    SettingsCharm.Margin = settingsDas2;

                    var searchDas3 = new Thickness(1, 38, 13.141, -44.89);
                    SearchText.Margin = searchDas3;

                    var shareDas3 = new Thickness(0, 59, 12, 0);
                    ShareText.Margin = shareDas3;

                    var winDas3 = new Thickness(0, -15, 12, 0);
                    WinText.Margin = winDas3;

                    var deviceDas3 = new Thickness(0, 7, 12, 0);
                    DevicesText.Margin = deviceDas3;

                    var settingsDas3 = new Thickness(0, 59, 12, -63);
                    SettingsText.Margin = settingsDas3;

                    var searchDas4 = new Thickness(0, -25, 0, 0);
                    SearchHover.Margin = searchDas4;

                    var shareDas4 = new Thickness(0, 0, 0, -25);
                    ShareHover.Margin = shareDas4;

                    var winDas4 = new Thickness(0, 25, 0, -50);
                    WinHover.Margin = winDas4;

                    var deviceDas4 = new Thickness(0, 50, 0, -75);
                    DevicesHover.Margin = deviceDas4;

                    var settingsDas4 = new Thickness(0, 75, 0, -100);
                    SettingsHover.Margin = settingsDas4;
                }

                this.Left = dispWidth - 86;

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
                    if (numVal < dispWidth - 86)
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
                        WinCharm.Visibility = Visibility.Hidden;
                        WinCharmHover.Visibility = Visibility.Visible;
                        WinHover.Visibility = Visibility.Visible;
                    }

                    if (mouseIn == false && keyboardShortcut == true && activeIcon != 2)
                    {
                        WinCharmUse = false;
                        WinCharm.Visibility = Visibility.Visible;
                        WinCharmHover.Visibility = Visibility.Hidden;
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
                        WinCharm.Visibility = Visibility.Visible;
                        WinCharmHover.Visibility = Visibility.Hidden;
                        WinHover.Visibility = Visibility.Hidden;
                        DevicesHover.Visibility = Visibility.Hidden;
                        SettingsHover.Visibility = Visibility.Hidden;
                    }
                }

                if (Keyboard.IsKeyDown(Key.Down) && swipeIn == false && keyboardTimer < 1 && waitTimer < 10 || Keyboard.IsKeyDown(Key.Tab) && swipeIn == false && keyboardTimer < 1 && waitTimer < 10)
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
                        WinCharm.Visibility = Visibility.Hidden;
                        WinCharmHover.Visibility = Visibility.Visible;
                        WinHover.Visibility = Visibility.Visible;
                    }

                    if (mouseIn == false && keyboardShortcut == true && activeIcon != 2)
                    {
                        WinCharmUse = false;
                        WinCharm.Visibility = Visibility.Visible;
                        WinCharmHover.Visibility = Visibility.Hidden;
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
                        WinCharm.Visibility = Visibility.Visible;
                        WinCharmHover.Visibility = Visibility.Hidden;
                        WinHover.Visibility = Visibility.Hidden;
                        DevicesHover.Visibility = Visibility.Hidden;
                        SettingsHover.Visibility = Visibility.Hidden;
                    }
                    }

                if (Keyboard.IsKeyDown(Key.Up) && swipeIn == false && keyboardTimer < 1 && waitTimer > 40)
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
                        WinCharm.Visibility = Visibility.Hidden;
                        WinCharmHover.Visibility = Visibility.Visible;
                        WinHover.Visibility = Visibility.Visible;
                    }

                    if (mouseIn == false && keyboardShortcut == true && activeIcon != 2)
                    {
                        WinCharmUse = false;
                        WinCharm.Visibility = Visibility.Visible;
                        WinCharmHover.Visibility = Visibility.Hidden;
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
                        WinCharm.Visibility = Visibility.Visible;
                        WinCharmHover.Visibility = Visibility.Hidden;
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
                        WinCharm.Visibility = Visibility.Hidden;
                        WinCharmHover.Visibility = Visibility.Visible;
                        WinHover.Visibility = Visibility.Visible;
                    }

                    if (mouseIn == false && keyboardShortcut == true && activeIcon != 2)
                    {
                        WinCharmUse = false;
                        WinCharm.Visibility = Visibility.Visible;
                        WinCharmHover.Visibility = Visibility.Hidden;
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
                        WinCharm.Visibility = Visibility.Visible;
                        WinCharmHover.Visibility = Visibility.Hidden;
                        WinHover.Visibility = Visibility.Hidden;
                        DevicesHover.Visibility = Visibility.Hidden;
                        SettingsHover.Visibility = Visibility.Hidden;
                    }
                }

                if (Keyboard.IsKeyDown(Key.Enter) && keyboardShortcut == true || Keyboard.IsKeyDown(Key.Space) && keyboardShortcut == true)
                {
                    if (activeIcon == 0)
                    {
                            byte winKey = (byte)KeyInterop.VirtualKeyFromKey(Key.LWin);
                            byte sKey = (byte)KeyInterop.VirtualKeyFromKey(Key.S);
                            const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
                            const uint KEYEVENTF_KEYUP = 0x0002;
                            charmsUse = false;
                            _ = keybd_event(winKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
                            _ = keybd_event(sKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
                            _ = keybd_event(winKey, 0, KEYEVENTF_KEYUP, 0);
                            _ = keybd_event(sKey, 0, KEYEVENTF_KEYUP, 0);

                            swipeIn = false;
                            keyboardShortcut = false;
                            charmsAppear = false;
                            charmsUse = false;

                            if (useAnimations == false)
                            {
                                this.Opacity = 0.001;
                                CharmsClock.Opacity = 0.001;

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

                        if (activeIcon == 1)
                        {
                            byte printScreenKey = (byte)KeyInterop.VirtualKeyFromKey(Key.PrintScreen);
                            const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
                            const uint KEYEVENTF_KEYUP = 0x0002;
                            charmsUse = false;
                            _ = keybd_event(printScreenKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
                            _ = keybd_event(printScreenKey, 0, KEYEVENTF_KEYUP, 0);

                            swipeIn = false;
                            keyboardShortcut = false;
                            charmsAppear = false;
                            charmsUse = false;
                            if (useAnimations == false)
                            {
                                this.Opacity = 0.001;
                                CharmsClock.Opacity = 0.001;

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

                        if (activeIcon == 2)
                        {
                            byte winKey = (byte)KeyInterop.VirtualKeyFromKey(Key.LWin);
                            const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
                            const uint KEYEVENTF_KEYUP = 0x0002;
                            charmsUse = false;
                            _ = keybd_event(winKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
                            _ = keybd_event(winKey, 0, KEYEVENTF_KEYUP, 0);

                            swipeIn = false;
                            keyboardShortcut = false;
                            charmsAppear = false;
                            charmsUse = false;
                            if (useAnimations == false)
                            {
                                this.Opacity = 0.001;
                                CharmsClock.Opacity = 0.001;

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

                        if (activeIcon == 3)
                        {
                            byte winKey = (byte)KeyInterop.VirtualKeyFromKey(Key.LWin);
                            byte pKey = (byte)KeyInterop.VirtualKeyFromKey(Key.P);
                            const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
                            const uint KEYEVENTF_KEYUP = 0x0002;
                            charmsUse = false;
                            _ = keybd_event(winKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
                            _ = keybd_event(pKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
                            _ = keybd_event(winKey, 0, KEYEVENTF_KEYUP, 0);
                            _ = keybd_event(pKey, 0, KEYEVENTF_KEYUP, 0);

                            swipeIn = false;
                            keyboardShortcut = false;
                            charmsAppear = false;
                            charmsUse = false;
                            if (useAnimations == false)
                            {
                                this.Opacity = 0.001;
                                CharmsClock.Opacity = 0.001;

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

                        if (activeIcon == 4)
                        {
                            byte winKey = (byte)KeyInterop.VirtualKeyFromKey(Key.LWin);
                            byte iKey = (byte)KeyInterop.VirtualKeyFromKey(Key.I);
                            const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
                            const uint KEYEVENTF_KEYUP = 0x0002;
                            charmsUse = false;
                            _ = keybd_event(winKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
                            _ = keybd_event(iKey, 0, KEYEVENTF_EXTENDEDKEY, 0);
                            _ = keybd_event(winKey, 0, KEYEVENTF_KEYUP, 0);
                            _ = keybd_event(iKey, 0, KEYEVENTF_KEYUP, 0);

                            swipeIn = false;
                            keyboardShortcut = false;
                            charmsAppear = false;
                            charmsUse = false;
                            if (useAnimations == false)
                            {
                                this.Opacity = 0.001;
                                CharmsClock.Opacity = 0.001;

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
                            waitTimer = 0;
                }

                if (WinDown.Visibility == Visibility.Visible && WinHover.Visibility == Visibility.Visible)
                {
                    var brush = (Brush)converter.ConvertFromString("#444444");
                    FadeBlocker.Background = brush;
                }

                if (WinHover.Visibility == Visibility.Visible && WinDown.Visibility != Visibility.Visible)
                {
                    WinCharmDown.Visibility = Visibility.Hidden;
                    var brush = (Brush)converter.ConvertFromString("#333333");
                    FadeBlocker.Background = brush;
                }

                if (WinHover.Visibility != Visibility.Visible && WinDown.Visibility != Visibility.Visible)
                {
                    WinCharmDown.Visibility = Visibility.Hidden;
                    var brush = (Brush)converter.ConvertFromString("#111111");
                    FadeBlocker.Background = brush;
                }

                if (charmsUse == false)
                {
                    SearchHover.Visibility = Visibility.Hidden;
                    ShareHover.Visibility = Visibility.Hidden;

                    WinCharm.Visibility = Visibility.Visible;
                    WinCharmHover.Visibility = Visibility.Hidden;
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
                    WinCharmHover.Visibility = Visibility.Hidden;
                    WinHover.Visibility = Visibility.Hidden;
                    DevicesHover.Visibility = Visibility.Hidden;
                    SettingsHover.Visibility = Visibility.Hidden;
                    activeIcon = 6;
                }

                if (keyboardShortcut == false)
                {
                    if (numVal > dispWidth - 12 & numVal2 < 12 && noTopRight == false || numVal > dispWidth - 12 & numVal2 > dispHeight - 12 && noBottomRight == false)
                    {
                        swipeIn = true;
                        charmsTimer += 1;
                    }

                    if (numVal > dispWidth - 12 & numVal2 < 12 && noTopRight == false || numVal > dispWidth - 12 & numVal2 > dispHeight - 12 && noBottomRight == false)
                    {
                        charmsWait += 1;
                    }

                    if (numVal < dispWidth - 86 && swipeIn == true && useAnimations == true || numVal > dispWidth - 85 && outofTime == true)
                    {
                        swipeIn = false;
                    }

                    if (numVal < dispWidth - 86 && useAnimations == false)
                    {
                        charmsTimer = 0;
                        charmsAppear = false;
                        charmsUse = false;
                        mouseIn = false;

                        this.Opacity = 0.001;
                        CharmsClock.Opacity = 0.001;

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
                        DevicesCharmInactive.Visibility = Visibility.Visible;
                        SettingsCharmInactive.Visibility = Visibility.Visible;
                    }

                    if (charmsAppear == true)
                    {
                        charmsTimer = 1945;
                    }
                }

                if (numVal < dispWidth - 86 && swipeIn == true && keyboardShortcut == false)
                {

        charmsAppear = false;
                    charmsUse = false;
                    if (useAnimations == false)
                    {
                        this.Opacity = 0.001;
                        CharmsClock.Opacity = 0.001;

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
                    DevicesCharmInactive.Visibility = Visibility.Visible;
                    SettingsCharmInactive.Visibility = Visibility.Visible;
                }

                if (numVal > dispWidth - 85 & numVal2 < 12 && charmsTimer > 100 && keyboardShortcut == false && outofTime == false || numVal > dispWidth - 85 & numVal2 > dispHeight - 12 && charmsTimer > 100 && keyboardShortcut == false || charmsAppear == true && keyboardShortcut == false || keyboardShortcut == true)
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
                }
                    charmsAppear = true;

                    if (charmsAppear == true && charmsUse == false)
                    {
                        SearchCharmInactive.Visibility = Visibility.Visible;
                        ShareCharmInactive.Visibility = Visibility.Visible;
                        NoColor.Visibility = Visibility.Visible;
                        WinCharmInactive.Visibility = Visibility.Visible;
                        DevicesCharmInactive.Visibility = Visibility.Visible;
                        SettingsCharmInactive.Visibility = Visibility.Visible;
                    }

                    if (charmsAppear == true && charmsUse == true)
                    {
                        SearchCharmInactive.Visibility = Visibility.Hidden;
                        ShareCharmInactive.Visibility = Visibility.Hidden;
                        NoColor.Visibility = Visibility.Hidden;
                        WinCharmInactive.Visibility = Visibility.Hidden;
                        DevicesCharmInactive.Visibility = Visibility.Hidden;
                        SettingsCharmInactive.Visibility = Visibility.Hidden;
                    }

                } else
                {
                    if (charmsAppear == false)
                    {
                        SearchCharmInactive.Visibility = Visibility.Hidden;
                        ShareCharmInactive.Visibility = Visibility.Hidden;
                        NoColor.Visibility = Visibility.Hidden;
                        WinCharmInactive.Visibility = Visibility.Hidden;
                        DevicesCharmInactive.Visibility = Visibility.Hidden;
                        SettingsCharmInactive.Visibility = Visibility.Hidden;
                    }

                    if (useAnimations == false)
                    {
                        this.Opacity = 0.001;
                    }
                }

                if (numVal > dispWidth - 85 & numVal2 > 208 & numVal2 < dispHeight - 702 && keyboardShortcut == false && swipeIn == true && useAnimations == true && outofTime == false || numVal > dispWidth - 85 & numVal2 > 193 & numVal2 < dispHeight - 202 && keyboardShortcut == false && swipeIn == true && useAnimations == true && outofTime == false || numVal > dispWidth - 85 & numVal2 > 208 & numVal2 < dispHeight - 702 && keyboardShortcut == false && useAnimations == false && outofTime == false || numVal > dispWidth - 85 & numVal2 > 193 & numVal2 < dispHeight - 202 && keyboardShortcut == false && useAnimations == false && outofTime == false || Keyboard.IsKeyDown(Key.LWin) && Keyboard.IsKeyDown(Key.C))
                {
                    this.Focus();
                    this.Activate();

                    if (charmsAppear == false)
                    {
                        charmsAppear = true;
                    }

                    if (keyboardShortcut == true && useAnimations == true)
                    {
                        MetroColor.Opacity = 1.0;
                        WinCharm.Opacity = 1.0;
                        CharmBG.Opacity = 0.001;
                    }

                    if (useAnimations == false)
                    {
                        var brush = (Brush)converter.ConvertFromString("#ff111111");
                        Background = brush;
                        this.Opacity = 1.0;
                    }

                    if (scrollSearch != 0 && scrollSearch < 199 && useAnimations == true)
                    {
                        var brush = (Brush)converter.ConvertFromString("#ff111111");
                        Background = brush;
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
                    if (WinCharmUse == false)
                    {
                        WinCharm.Visibility = Visibility.Visible;
                        WinCharmHover.Visibility = Visibility.Hidden;
                    }

                    if (WinCharmUse == true)
                    {
                        WinCharm.Visibility = Visibility.Hidden;
                        WinCharmHover.Visibility = Visibility.Visible;
                    }
                    DevicesCharm.Visibility = Visibility.Visible;
                    SettingsCharm.Visibility = Visibility.Visible;

                    SearchCharmInactive.Visibility = Visibility.Hidden;
                    ShareCharmInactive.Visibility = Visibility.Hidden;
                    WinCharmInactive.Visibility = Visibility.Hidden;
                    NoColor.Visibility = Visibility.Hidden;
                    DevicesCharmInactive.Visibility = Visibility.Hidden;
                    SettingsCharmInactive.Visibility = Visibility.Hidden;

                    if (Keyboard.IsKeyDown(Key.LWin) && Keyboard.IsKeyDown(Key.C))
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
                            WinCharm.Visibility = Visibility.Hidden;
                            WinCharmHover.Visibility = Visibility.Visible;
                            WinHover.Visibility = Visibility.Visible;
                        }

                        if (mouseIn == false && keyboardShortcut == true && activeIcon != 2)
                        {
                            WinCharmUse = false;
                            WinCharm.Visibility = Visibility.Visible;
                            WinCharmHover.Visibility = Visibility.Hidden;
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

                        if (useAnimations == false)
                        {
                            this.Opacity = 1.0;
                            CharmsClock.Show();
                            CharmsClock.Opacity = 1.0;
                        }

                        if (useAnimations == true)
                        {
                            this.Opacity = IHOb;
                            CharmsClock.Show();
                            CharmsClock.Opacity = IHOb;
                        }

                        keyboardShortcut = true;
                    } else
                    {
                        if (useAnimations == false)
                        {
                            this.Opacity = 1.0;
                            CharmsClock.Show();
                            CharmsClock.Opacity = 1.0;
                        }

                        if (useAnimations == true)
                        {
                            this.Opacity = IHOb;
                            CharmsClock.Show();
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

                        this.Opacity = 0.001;
                        CharmsClock.Opacity = 0.001;

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
                        DevicesCharmInactive.Visibility = Visibility.Visible;
                        SettingsCharmInactive.Visibility = Visibility.Visible;
                    }
                    else
                    {
                        escKey = true;
                    }
                }

                if (SystemParameters.HighContrast == false)
                {
                    CharmBorder.Visibility = Visibility.Hidden;
                    CharmBorder.Background = (Brush)converter.ConvertFromString("#111111");
                    SearchText.Foreground = (Brush)converter.ConvertFromString("#a0a0a0");
                    ShareText.Foreground = (Brush)converter.ConvertFromString("#a0a0a0");
                    WinText.Foreground = (Brush)converter.ConvertFromString("#a0a0a0");
                    DevicesText.Foreground = (Brush)converter.ConvertFromString("#a0a0a0");
                    SettingsText.Foreground = (Brush)converter.ConvertFromString("#a0a0a0");
                    CharmBG.Background = (Brush)converter.ConvertFromString("#111111");
                    if (charmsUse == true)
                    {
                        HighColor.Visibility = Visibility.Hidden;
                        MetroColor.Visibility = Visibility.Visible;
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
                    CharmBorder.Visibility = Visibility.Visible;
                    CharmBorder.Background = SystemColors.WindowTextBrush;
                    SearchText.Foreground = SystemColors.WindowTextBrush;
                    ShareText.Foreground = SystemColors.WindowTextBrush;
                    WinText.Foreground = SystemColors.WindowTextBrush;
                    DevicesText.Foreground = SystemColors.WindowTextBrush;
                    SettingsText.Foreground = SystemColors.WindowTextBrush;
                    CharmBG.Background = SystemColors.MenuBrush;
                    if (charmsUse == true)
                    {
                        MetroColor.Visibility = Visibility.Hidden;
                        HighColor.Visibility = Visibility.Visible;
                    }
                    SearchHover.Background = SystemColors.MenuHighlightBrush;
                    ShareHover.Background = SystemColors.MenuHighlightBrush;
                    WinHover.Background = SystemColors.MenuHighlightBrush;
                    DevicesHover.Background = SystemColors.MenuHighlightBrush;
                    SettingsHover.Background = SystemColors.MenuHighlightBrush;
                    SearchDown.Background = SystemColors.MenuBrush;
                    ShareDown.Background = SystemColors.MenuBrush;
                    WinDown.Background = SystemColors.MenuBrush;
                    DevicesDown.Background = SystemColors.MenuBrush;
                    SettingsDown.Background = SystemColors.MenuBrush;
                }
                CharmBorder.Opacity = CharmBG.Opacity;
            }));
        }

        private void Charms_MouseMove(object sender, System.EventArgs e)
        {
            mouseIn = true;
            twoInputs = true;
            if (this.IsActive == false && charmsUse == false && keyboardShortcut == false)
            {
                if (useAnimations == false)
                {
                    this.Opacity = 1.0;
                    CharmsClock.Opacity = 0.001;

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
                    CharmsClock.Opacity = 0.001;
                }
            }

            if (this.IsActive == true && charmsUse == true && keyboardShortcut == false)
            {
                this.Focus();
                this.Activate();
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
                    this.Opacity = 0.001;
                    CharmsClock.Opacity = 0.001;
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
            SearchDown.Visibility = Visibility.Visible;
            SearchHover.Visibility = Visibility.Hidden;
        }

        private void Share_MouseDown(object sender, System.EventArgs e)
        {
            ShareDown.Visibility = Visibility.Visible;
            ShareHover.Visibility = Visibility.Hidden;
        }

        private void Win_MouseDown(object sender, System.EventArgs e)
        {
            WinCharmDown.Visibility = Visibility.Visible;
            WinCharmHover.Visibility = Visibility.Hidden;
            WinDown.Visibility = Visibility.Visible;
            WinHover.Visibility = Visibility.Hidden;
        }

        private void Devices_MouseDown(object sender, System.EventArgs e)
        {
            DevicesDown.Visibility = Visibility.Visible;
            DevicesHover.Visibility = Visibility.Hidden;
        }

        private void Settings_MouseDown(object sender, System.EventArgs e)
        {
            SettingsDown.Visibility = Visibility.Visible;
            SettingsHover.Visibility = Visibility.Hidden;
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
        [DllImport("user32.dll", EntryPoint = "GetKeyboardState", SetLastError = true)]
        private static extern bool NativeGetKeyboardState([Out] byte[] keyStates);

        private static bool GetKeyboardState(byte[] keyStates)
        {
            if (keyStates == null)
                throw new ArgumentNullException("keyState");
            if (keyStates.Length != 256)
                throw new ArgumentException("The buffer must be 256 bytes long.", "keyState");
            return NativeGetKeyboardState(keyStates);
        }

        private static byte[] GetKeyboardState()
        {
            byte[] keyStates = new byte[256];
            if (!GetKeyboardState(keyStates))
                throw new Win32Exception(Marshal.GetLastWin32Error());
            return keyStates;
        }

        private static bool AnyKeyPressed()
        {
            byte[] keyState = GetKeyboardState();
            // skip the mouse buttons
            return keyState.Skip(8).Any(state => (state & 0x80) != 0);
        }

        public static Point ElementPointToScreenPoint(UIElement element, Point pointOnElement)
        {
            return element.PointToScreen(pointOnElement);
        }

        private void Search_MouseEnter(object sender, System.EventArgs e)
        {
            mouseIn = true;
            twoInputs = true;
            if (charmsAppear == true)
            {
                SearchHover.Visibility = Visibility.Visible;
            }
        }

        private void Search_MouseLeave(object sender, System.EventArgs e)
        {
            mouseIn = true;
            twoInputs = true;
            if (charmsAppear == true)
            {
                SearchHover.Visibility = Visibility.Hidden;
            }
        }

        private void Share_MouseEnter(object sender, System.EventArgs e)
        {
            mouseIn = true;
            twoInputs = true;
            if (charmsAppear == true)
            {
                ShareHover.Visibility = Visibility.Visible;
            }
        }

        private void Share_MouseLeave(object sender, System.EventArgs e)
        {
            mouseIn = true;
            twoInputs = true;
            if (charmsAppear == true)
            {
                ShareHover.Visibility = Visibility.Hidden;
            }
        }

        private void Win_MouseEnter(object sender, System.EventArgs e)
        {
            mouseIn = true;
            twoInputs = true;
            WinCharmUse = true;
            if (charmsAppear == true)
            {
                WinCharm.Visibility = Visibility.Hidden;
                WinCharmHover.Visibility = Visibility.Visible;
                WinHover.Visibility = Visibility.Visible;
            }
        }

        private void Win_MouseLeave(object sender, System.EventArgs e)
        {
            mouseIn = true;
            twoInputs = true;
            WinCharmUse = false;
            if (charmsAppear == true)
            {
                WinCharm.Visibility = Visibility.Visible;
                WinCharmHover.Visibility = Visibility.Hidden;
                WinHover.Visibility = Visibility.Hidden;
            }
        }

        private void Devices_MouseEnter(object sender, System.EventArgs e)
        {
            mouseIn = true;
            twoInputs = true;
            if (charmsAppear == true)
            {
                DevicesHover.Visibility = Visibility.Visible;
            }
        }

        private void Devices_MouseLeave(object sender, System.EventArgs e)
        {
            mouseIn = true;
            twoInputs = true;
            if (charmsAppear == true)
            {
                DevicesHover.Visibility = Visibility.Hidden;
            }
        }

        private void Settings_MouseEnter(object sender, System.EventArgs e)
        {
            mouseIn = true;
            twoInputs = true;
            SettingsHover.Visibility = Visibility.Visible;
        }

        private void Settings_MouseLeave(object sender, System.EventArgs e)
        {
            mouseIn = true;
            twoInputs = true;
            SettingsHover.Visibility = Visibility.Hidden;
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
                DevicesCharmInactive.Visibility = Visibility.Visible;
                SettingsCharmInactive.Visibility = Visibility.Visible;

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
                    this.Opacity = 0.001;
                    CharmsClock.Opacity = 0.001;
                }
            }

            if (this.IsActive == false && charmsUse == false && twoInputs == false)
            {
                if (useAnimations == false)
                {
                    this.Opacity = 0.001;
                    CharmsClock.Opacity = 0.001;
                }
            }

            if (this.IsActive == false && charmsUse == true && twoInputs == false)
            {
                charmsUse = false;
                if (useAnimations == false)
                {
                    this.Opacity = 0.001;
                    CharmsClock.Opacity = 0.001;
                }
            }

            if (this.IsActive == false && charmsUse == true && twoInputs == true)
            {
                charmsUse = false;
                if (useAnimations == false)
                {
                    this.Opacity = 0.001;
                    CharmsClock.Opacity = 0.001;
                }
            }

            charmsTimer = 0;
        }
    }
}
