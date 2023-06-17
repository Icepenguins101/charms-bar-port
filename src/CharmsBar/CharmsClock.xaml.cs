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
using System.Reflection.Emit;
using System.ComponentModel;
using System.Threading.Tasks;
using Windows.Networking.Connectivity;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using System.Windows.Forms;
using System.Reflection;
using Microsoft.Win32;
using System.Collections.ObjectModel;
using static System.Resources.ResXFileRef;
using System.Net.NetworkInformation;
using System.Threading;

namespace CharmsBarPort
{
    public partial class CharmsClock : Window
    {
        public Microsoft.Win32.RegistryKey localKey = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, RegistryView.Registry64);
        public int isAirPlaneOn = 0;
        public BackgroundWorker CheckSignal = new BackgroundWorker();
        public string nw4 = "";
        public string nw5 = "";
        public string isDark = "";
        public string hasDrivers = "";
        public string isEthernet = "";
        public CharmsClock()
        {
            var dispWidth = SystemParameters.PrimaryScreenWidth;
            var dispHeight = SystemParameters.PrimaryScreenHeight;
            Topmost = true;
            ShowInTaskbar = false;
            WindowStyle = WindowStyle.None;
            ResizeMode = ResizeMode.NoResize;
            AllowsTransparency = true;
            Height = 140;
            WindowStartupLocation = WindowStartupLocation.Manual;
            Left = 51;
            Top = dispHeight - 188;
            BrushConverter converter = new();
            var brush = (Brush)converter.ConvertFromString("#f0111111");
            Background = brush;
            CheckSignal.DoWork += CheckSignal_DoWork;
            CheckSignal.ProgressChanged += CheckSignal_ProgressChanged;
            CheckSignal.WorkerReportsProgress = true;
            System.Windows.Forms.Application.ThreadException += new ThreadExceptionEventHandler(CharmsClock.Form1_UIThreadException);
            InitializeComponent();
            _initTimer();
        }

        private System.Windows.Forms.Timer t = null;
        private readonly Dispatcher dispatcher = Dispatcher.CurrentDispatcher;

        private void _initTimer()
        {
            System.Windows.Forms.Timer t = new System.Windows.Forms.Timer();
            t.Interval = 1;
            t.Tick += OnTimedEvent;
            t.Enabled = true;
            t.Start();
        }

        private void OnTimedEvent(object sender, EventArgs e)
        {
            dispatcher.BeginInvoke((Action)(() =>
            {
                try
                {
                    RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\ImmersiveShell\\EdgeUi", false);
                    if (key != null)
                    {
                        // get value 
                        string noClock = key.GetValue("DisableCharmsClock", -1, RegistryValueOptions.None).ToString(); //this is not in Windows 8.1, but used to remove the Charms Clock

                            if (noClock == "-1")
                            {
                                noClocks.Content = "0";
                            }
                            else
                            {
                                noClocks.Content = noClock;
                            }

                        }
                        key.Close();
                    }
                
                catch (Exception ex)  //just for demonstration...it's always best to handle specific exceptions
                {
                    //react appropriately
                }

                if (noClocks.Content == "-1" || noClocks.Content == "0")
                {

                if (SystemParameters.HighContrast == false)
                {
                    isDark = "";
                }

                if (SystemParameters.HighContrast == true)
                {
                    System.Drawing.Color col = System.Drawing.ColorTranslator.FromHtml(SystemColors.WindowBrush.ToString());
                    if (col.R * 0.2126 + col.G * 0.7152 + col.B * 0.0722 < 255 / 2)
                    {
                        // dark color
                        isDark = "";
                    }
                    else
                    {
                        // light color
                        isDark = "Dark";
                    }
                }

                Date.Content = DateTime.Today.ToString("MMMM d");
                if (DateTime.Today.ToString("dddd") != "Sunday" && DateTime.Today.ToString("dddd") != "Monday" && DateTime.Today.ToString("dddd") != "Friday")
                {
                    Week.Content = DateTime.Today.ToString("dddd  ");
                }
                else
                {
                    Week.Content = DateTime.Today.ToString("dddd      ");
                }
                Clocks.Content = DateTime.Now.ToString("h ");
                Clocked.Content = DateTime.Now.ToString("mm");

                if (SystemParameters.HighContrast == false)
                {
                    ClockBorder.Visibility = Visibility.Hidden;
                    BrushConverter converter = new();
                    ClockLines.Foreground = (Brush)converter.ConvertFromString("#ffffff");
                    Clocks.Foreground = (Brush)converter.ConvertFromString("#ffffff");
                    Week.Foreground = (Brush)converter.ConvertFromString("#ffffff");
                    Date.Foreground = (Brush)converter.ConvertFromString("#ffffff");
                    Clocked.Foreground = (Brush)converter.ConvertFromString("#ffffff");
                }

                if (SystemParameters.HighContrast == true)
                {
                    ClockBorder.Visibility = Visibility.Visible;
                    ClockLines.Foreground = SystemColors.WindowTextBrush;
                    Clocks.Foreground = SystemColors.WindowTextBrush;
                    Week.Foreground = SystemColors.WindowTextBrush;
                    Date.Foreground = SystemColors.WindowTextBrush;
                    Clocked.Foreground = SystemColors.WindowTextBrush;
                }

                try
                {
                    while (!CheckSignal.IsBusy)
                    {
                        CheckSignal.RunWorkerAsync();
                    }
                }

                catch(Exception err)
                {
                    
                }

                CheckBatteryStatus();
                ClockBorder.BorderBrush = SystemColors.WindowTextBrush;
                var localKey = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, RegistryView.Registry64);
                var localKey2 = localKey.OpenSubKey("SYSTEM\\ControlSet001\\Control\\RadioManagement\\SystemRadioState");
                var isAirPlaneOn = localKey2.GetValue("", "").ToString();
                ClockBorder.Background = SystemColors.WindowBrush;
                var nw = "";
                var nw2 = "";
                var nw3 = "";

                nw = IsConnected().ToString();
                nw2 = IsLocal().ToString();
                nw3 = IsWeak().ToString();

                if (SystemParameters.HighContrast == false)
                {
                    NoDrivers.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon103.png", UriKind.Relative));
                    NoInternet.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon115.png", UriKind.Relative));
                    Ethernet.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon106.png", UriKind.Relative));
                    NoInternetFound.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon112.png", UriKind.Relative));
                    IsCharging.Source = new BitmapImage(new Uri(@"/Assets/Images/BatteryFullCharging.png", UriKind.Relative));
                    Airplane.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon118.png", UriKind.Relative));
                }

                if (SystemParameters.HighContrast == true)
                {
                    NoDrivers.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon103" + isDark + ".png", UriKind.Relative));
                    NoInternet.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon115" + isDark + ".png", UriKind.Relative));
                    Ethernet.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon106" + isDark + ".png", UriKind.Relative));
                    NoInternetFound.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon112" + isDark + ".png", UriKind.Relative));
                    IsCharging.Source = new BitmapImage(new Uri(@"/Assets/Images/BatteryFullCharging" + isDark + ".png", UriKind.Relative));
                    Airplane.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon118" + isDark + ".png", UriKind.Relative));
                }


                if (isEthernet.IndexOf("Ethernet") != -1)
                {
                    NoDrivers.Visibility = Visibility.Hidden;
                    NoInternet.Visibility = Visibility.Hidden;
                    NoInternetFound.Visibility = Visibility.Hidden;
                    Ethernet.Visibility = Visibility.Visible;
                    HasInternet.Visibility = Visibility.Hidden;
                    WeakInternet.Visibility = Visibility.Hidden;
                    Airplane.Visibility = Visibility.Hidden;
                }

                if (hasDrivers.IndexOf("Thereisno") != -1)
                {
                    NoDrivers.Visibility = Visibility.Visible;
                    NoInternet.Visibility = Visibility.Hidden;
                    NoInternetFound.Visibility = Visibility.Hidden;
                    Ethernet.Visibility = Visibility.Hidden;
                    HasInternet.Visibility = Visibility.Hidden;
                    WeakInternet.Visibility = Visibility.Hidden;
                    Airplane.Visibility = Visibility.Hidden;
                }
                else
                {
                    if (SystemParameters.HighContrast == false)
                    {
                        if (nw == "True" && nw2 == "False" && nw3 == "False" && isAirPlaneOn == "0")
                        {
                            NoDrivers.Visibility = Visibility.Hidden;
                            NoInternet.Visibility = Visibility.Hidden;
                            NoInternetFound.Visibility = Visibility.Hidden;
                            Ethernet.Visibility = Visibility.Hidden;
                            HasInternet.Visibility = Visibility.Visible;
                            WeakInternet.Visibility = Visibility.Hidden;
                            Airplane.Visibility = Visibility.Hidden;

                            if (nw4.StartsWith("100") == true || nw4.StartsWith("9") == true || nw4.StartsWith("8") == true)
                            {
                                HasInternet.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon151.png", UriKind.Relative));
                                WeakInternet.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon133.png", UriKind.Relative));
                            }

                            if (nw4.StartsWith("6") == true || nw4.StartsWith("7") == true)
                            {
                                HasInternet.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon148.png", UriKind.Relative));
                                WeakInternet.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon130.png", UriKind.Relative));
                            }

                            if (nw4.StartsWith("4") == true || nw4.StartsWith("5") == true)
                            {
                                HasInternet.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon145.png", UriKind.Relative));
                                WeakInternet.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon127.png", UriKind.Relative));
                            }

                            if (nw4.StartsWith("2") == true || nw4.StartsWith("3") == true)
                            {
                                HasInternet.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon142.png", UriKind.Relative));
                                WeakInternet.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon124.png", UriKind.Relative));
                            }

                            if (nw4.StartsWith("0") == true || nw4.StartsWith("1") == true)
                            {
                                HasInternet.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon139.png", UriKind.Relative));
                                WeakInternet.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon121.png", UriKind.Relative));
                            }
                        }
                    }

                    if (SystemParameters.HighContrast == true)
                    {
                        if (nw == "True" && nw2 == "False" && nw3 == "False" && isAirPlaneOn == "0")
                        {
                            NoDrivers.Visibility = Visibility.Hidden;
                            NoInternet.Visibility = Visibility.Hidden;
                            NoInternetFound.Visibility = Visibility.Hidden;
                            Ethernet.Visibility = Visibility.Hidden;
                            HasInternet.Visibility = Visibility.Visible;
                            WeakInternet.Visibility = Visibility.Hidden;
                            Airplane.Visibility = Visibility.Hidden;

                            if (nw4.StartsWith("100") == true || nw4.StartsWith("9") == true || nw4.StartsWith("8") == true)
                            {
                                HasInternet.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon151" + isDark + ".png", UriKind.Relative));
                                WeakInternet.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon133" + isDark + ".png", UriKind.Relative));
                            }

                            if (nw4.StartsWith("6") == true || nw4.StartsWith("7") == true)
                            {
                                HasInternet.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon148" + isDark + ".png", UriKind.Relative));
                                WeakInternet.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon130" + isDark + ".png", UriKind.Relative));
                            }

                            if (nw4.StartsWith("4") == true || nw4.StartsWith("5") == true)
                            {
                                HasInternet.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon145" + isDark + ".png", UriKind.Relative));
                                WeakInternet.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon127" + isDark + ".png", UriKind.Relative));
                            }

                            if (nw4.StartsWith("2") == true || nw4.StartsWith("3") == true)
                            {
                                HasInternet.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon142" + isDark + ".png", UriKind.Relative));
                                WeakInternet.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon124" + isDark + ".png", UriKind.Relative));
                            }

                            if (nw4.StartsWith("0") == true || nw4.StartsWith("1") == true)
                            {
                                HasInternet.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon139" + isDark + ".png", UriKind.Relative));
                                WeakInternet.Source = new BitmapImage(new Uri(@"/Assets/Images/Icon121" + isDark + ".png", UriKind.Relative));
                            }
                        }
                    }

                    if (nw == "True" && nw2 == "False" && nw3 == "False" && isAirPlaneOn == "1")
                    {
                        NoDrivers.Visibility = Visibility.Hidden;
                        NoInternet.Visibility = Visibility.Hidden;
                        NoInternetFound.Visibility = Visibility.Hidden;
                        Ethernet.Visibility = Visibility.Hidden;
                        HasInternet.Visibility = Visibility.Hidden;
                        WeakInternet.Visibility = Visibility.Hidden;
                        Airplane.Visibility = Visibility.Visible;
                    }

                    if (nw2 == "True" && nw == "False" && nw3 == "False" && isAirPlaneOn == "0")
                    {
                        NoDrivers.Visibility = Visibility.Hidden;
                        NoInternet.Visibility = Visibility.Hidden;
                        NoInternetFound.Visibility = Visibility.Hidden;
                        Ethernet.Visibility = Visibility.Hidden;
                        HasInternet.Visibility = Visibility.Hidden;
                        WeakInternet.Visibility = Visibility.Visible;
                        Airplane.Visibility = Visibility.Hidden;
                    }

                    if (nw == "False" && nw2 == "False" && nw3 == "False" && isAirPlaneOn == "0" && nw5 == "SoftwareOn")
                    {
                        NoDrivers.Visibility = Visibility.Hidden;
                        NoInternet.Visibility = Visibility.Visible;
                        NoInternetFound.Visibility = Visibility.Hidden;
                        Ethernet.Visibility = Visibility.Hidden;
                        HasInternet.Visibility = Visibility.Hidden;
                        WeakInternet.Visibility = Visibility.Hidden;
                        Airplane.Visibility = Visibility.Hidden;
                    }

                    if (nw == "False" && nw2 == "False" && nw3 == "False" && isAirPlaneOn == "0" && nw5 == "SoftwareOff")
                    {
                        NoDrivers.Visibility = Visibility.Hidden;
                        NoInternet.Visibility = Visibility.Hidden;
                        NoInternetFound.Visibility = Visibility.Visible;
                        Ethernet.Visibility = Visibility.Hidden;
                        HasInternet.Visibility = Visibility.Hidden;
                        WeakInternet.Visibility = Visibility.Hidden;
                        Airplane.Visibility = Visibility.Hidden;
                    }

                    if (nw == "False" && nw2 == "False" && nw3 == "True" && isAirPlaneOn == "0")
                    {
                        NoDrivers.Visibility = Visibility.Hidden;
                        NoInternet.Visibility = Visibility.Hidden;
                        NoInternetFound.Visibility = Visibility.Hidden;
                        Ethernet.Visibility = Visibility.Hidden;
                        HasInternet.Visibility = Visibility.Hidden;
                        WeakInternet.Visibility = Visibility.Visible;
                        Airplane.Visibility = Visibility.Hidden;
                    }

                    if (nw == "False" && nw2 == "False" && nw3 == "False" && isAirPlaneOn == "1")
                    {
                        NoDrivers.Visibility = Visibility.Hidden;
                        NoInternet.Visibility = Visibility.Hidden;
                        NoInternetFound.Visibility = Visibility.Hidden;
                        Ethernet.Visibility = Visibility.Hidden;
                        HasInternet.Visibility = Visibility.Hidden;
                        WeakInternet.Visibility = Visibility.Hidden;
                        Airplane.Visibility = Visibility.Visible;
                    }
                }
                if (Clocks.Content.ToString().Length < 3 && Clocks.Content.ToString().StartsWith("1 ") == false && Clocks.Content.ToString().StartsWith("10") == false && Clocks.Content.ToString().StartsWith("11") == false && Clocks.Content.ToString().StartsWith("12") == false || Clocks.Content.ToString().Length == 2 && Clocks.Content.ToString().StartsWith("1 ") == false && Clocks.Content.ToString().StartsWith("10") == false && Clocks.Content.ToString().StartsWith("11") == false && Clocks.Content.ToString().StartsWith("12") == false)
                {
                    Clocks.Margin = new Thickness(94, 3, 0, -106);
                    ClockLines.Margin = new Thickness(138, -24.99, -190, -98);
                    Clocked.Margin = new Thickness(157, -17, -190, -198);
                    Week.Margin = new Thickness(267, 2, 0, -18);
                    Date.Margin = new Thickness(269, 3, 0, -24);
                }

                if (Clocks.Content.ToString().StartsWith("1 ") == true)
                {
                    Clocks.Margin = new Thickness(95, 3, 0, -106);
                    ClockLines.Margin = new Thickness(125, -24.99, -190, -98);
                    Clocked.Margin = new Thickness(144, -17, -190, -198);
                    Week.Margin = new Thickness(255, 2, 0, -18);
                    Date.Margin = new Thickness(255, 3, 0, -24);
                }

                if (Clocks.Content.ToString().StartsWith("10") == true)
                {
                    Clocks.Margin = new Thickness(95, 3, 0, -106);
                    ClockLines.Margin = new Thickness(169, -24.99, -190, -98);
                    Clocked.Margin = new Thickness(188, -17, -190, -198);
                    Week.Margin = new Thickness(298, 2, 0, -18);
                    Date.Margin = new Thickness(300, 4, 0, -24);
                }

                if (Clocks.Content.ToString().StartsWith("11") == true)
                {
                    Clocks.Margin = new Thickness(95, 3, 0, -106);
                    ClockLines.Margin = new Thickness(156, -24.99, -190, -98);
                    Clocked.Margin = new Thickness(174.5, -17, -190, -198);
                    Week.Margin = new Thickness(284, 2, 0, -18);
                    Date.Margin = new Thickness(284, 4, 0, -24);
                }

                if (Clocks.Content.ToString().StartsWith("12") == true)
                {
                    Clocks.Margin = new Thickness(95, 3, 0, -106);
                    ClockLines.Margin = new Thickness(169, -24.99, -190, -98);
                    Clocked.Margin = new Thickness(188, -17, -190, -198);
                    Week.Margin = new Thickness(298, 2, 0, -18);
                    Date.Margin = new Thickness(300, 4, 0, -24);
                }

                if (Clocks.Content.ToString().Length == 3 && Clocks.Content.ToString().StartsWith("10") == false && Clocks.Content.ToString().StartsWith("12") == false)
                {
                    Clocks.Margin = new Thickness(94, 3, 0, -106);
                    ClockLines.Margin = new Thickness(157, -24.99, -190, -98);
                    Clocked.Margin = new Thickness(175, -17, -190, -198);
                    Week.Margin = new Thickness(287, 2, 0, -18);
                    Date.Margin = new Thickness(287, 4, 0, -24);
                }

                if (Date.Content.ToString().Length > 6 || Date.Content.ToString().Length < 7 && Week.Content.ToString().Length > 9)
                {
                    if (Clocks.Content.ToString().Length == 2 && Clocks.Content.ToString().StartsWith("1") == false)
                    {
                        CharmClock.Margin = new Thickness(0, 0, 10, 0);
                        AutoResizer.Width = 391 + Date.Content.ToString().Length + 25;
                    }

                    if (Clocks.Content.ToString().Length == 2 && Clocks.Content.ToString().StartsWith("1") == true)
                    {
                        CharmClock.Margin = new Thickness(0, 0, 10, 0);
                        AutoResizer.Width = 391 + Date.Content.ToString().Length;
                    }

                    if (Clocks.Content.ToString().Length == 3 && Clocks.Content.ToString().StartsWith("10") == true)
                    {
                        CharmClock.Margin = new Thickness(0, 0, 10, 0);
                        AutoResizer.Width = 391 + Date.Content.ToString().Length + 63;
                    }

                    if (Clocks.Content.ToString().Length == 3 && Clocks.Content.ToString().StartsWith("12") == true)
                    {
                        CharmClock.Margin = new Thickness(0, 0, 10, 0);
                        AutoResizer.Width = 391 + Date.Content.ToString().Length + 63;
                    }

                    if (Clocks.Content.ToString().Length == 3 && Clocks.Content.ToString().StartsWith("11") == true)
                    {
                        CharmClock.Margin = new Thickness(0, 0, 10, 0);
                        AutoResizer.Width = 391 + Date.Content.ToString().Length + 63;
                    }

                    if (Clocks.Content.ToString().Length == 3 && Clocks.Content.ToString().StartsWith("10") == false && Clocks.Content.ToString().StartsWith("11") == false && Clocks.Content.ToString().StartsWith("12") == false && Clocks.Content.ToString().StartsWith("1") == false)
                    {
                        CharmClock.Margin = new Thickness(0, 0, 10, 0);
                        AutoResizer.Width = 391 + Date.Content.ToString().Length + 25;
                    }
                }

                ClockBorder.Width = this.Width;
                localKey.Close();
                    }
            }));
        }

        // Handle the UI exceptions by showing a dialog box, and asking the user whether
        // or not they wish to abort execution.
        private static void Form1_UIThreadException(object sender, ThreadExceptionEventArgs t)
        {
            System.Windows.Forms.Application.Restart();
        }

        private bool IsConnected()
        {
            try
            {
                var connectionProfile = NetworkInformation.GetInternetConnectionProfile();
                return (connectionProfile != null &&
                      (connectionProfile.GetNetworkConnectivityLevel() == NetworkConnectivityLevel.InternetAccess));
            }

            catch(Exception err)
            {
                System.Windows.Forms.Application.Restart();
                return false;
            }
        }
        private bool IsWeak()
        {
            var connectionProfile = NetworkInformation.GetInternetConnectionProfile();
            return (connectionProfile != null &&
                  (connectionProfile.GetNetworkConnectivityLevel() == NetworkConnectivityLevel.ConstrainedInternetAccess));
        }

        private bool IsLocal()
        {
            var connectionProfile = NetworkInformation.GetInternetConnectionProfile();
            return (connectionProfile != null &&
                  (connectionProfile.GetNetworkConnectivityLevel() == NetworkConnectivityLevel.LocalAccess));
        }

        private void CheckSignal_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                Process proc = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "netsh.exe",
                        Arguments = "wlan show interfaces",
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        CreateNoWindow = true
                    }
                };
                while (true)
                {
                    proc.Start();
                    string line;
                    int strength = 0;
                    string wifi;
                    while (!proc.StandardOutput.EndOfStream)
                    {
                        line = proc.StandardOutput.ReadLine();

                        if (line.Contains("Name"))
                        {
                            string tmpx = line.Split(':')[1].Split("%")[0];
                            isEthernet = tmpx.ToString();
                        }

                        if (line.Contains("There is"))
                        {
                            string tmp = line;
                            hasDrivers = tmp.Replace(" ", "");
                        }

                        if (line.Contains("Software"))
                        {
                            string tmp2 = line;
                            nw5 = tmp2.Replace(" ", "");
                        }

                        if (line.Contains("Signal"))
                        {
                            string tmp3 = line.Split(':')[1].Split("%")[0];
                            Int32.TryParse(tmp3, out strength);
                            nw4 = strength.ToString();
                            CheckSignal.ReportProgress(strength);
                        }
System.Threading.Thread.Sleep(10);
proc.WaitForExit(); //this should make it more stable, I hope
                    }
                }
            }

            catch(Exception ex)
            {
                
            }
        }

        static void CheckSignal_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            
        }

        private void CheckBatteryStatus()
        {
            var pw = SystemInformation.PowerStatus.BatteryChargeStatus.ToString();
            var pw2 = SystemInformation.PowerStatus.PowerLineStatus.ToString();
            double pw3 = SystemInformation.PowerStatus.BatteryLifePercent;
            var dasBoot = pw3.ToString();

            if (dasBoot.StartsWith("1") == true || dasBoot.StartsWith("0.9") == true && pw3 > 0.96 && dasBoot.StartsWith("1.0") == false)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/BatteryFull" + isDark + ".png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.9") == true && pw3 > 0.95 && pw3 < 0.96 && dasBoot.StartsWith("1") == false)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery90" + isDark + ".png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.9") == true && pw3 < 0.95)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery90" + isDark + ".png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.8") == true && pw3 > 0.86)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery80" + isDark + ".png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.8") == true && pw3 < 0.86)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery80" + isDark + ".png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.7") == true && pw3 > 0.76)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery70" + isDark + ".png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.7") == true && pw3 < 0.76)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery70" + isDark + ".png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.6") == true && pw3 > 0.66)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery60" + isDark + ".png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.6") == true && pw3 < 0.66)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery60" + isDark + ".png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.5") == true && pw3 > 0.56)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery50" + isDark + ".png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.5") == true && pw3 < 0.56)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery50" + isDark + ".png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.4") == true && pw3 > 0.46)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery40" + isDark + ".png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.4") == true && pw3 < 0.46)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery40" + isDark + ".png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.3") == true && pw3 > 0.36)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery30" + isDark + ".png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.3") == true && pw3 < 0.36)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery30" + isDark + ".png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.2") == true && pw3 > 0.26)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery20" + isDark + ".png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.2") == true && pw3 < 0.26)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery20" + isDark + ".png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.1") == true && pw3 > 0.16)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery10" + isDark + ".png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.1") == true && pw3 < 0.16)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery1" + isDark + ".png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.1") == true && pw3 > 0.6)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery5" + isDark + ".png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.1") == true && pw3 < 0.6)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery5" + isDark + ".png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.1") == true && pw3 > 0.2)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery1" + isDark + ".png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.1") == true && pw3 < 0.2)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery1" + isDark + ".png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.1") == true && pw3 < 0.1)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery0" + isDark + ".png", UriKind.Relative));
            }

            if (pw2 == "Online")
            {
                BatteryLife.Visibility = Visibility.Visible;
                IsCharging.Visibility = Visibility.Visible;
            }

            if (pw2 == "Offline")
            {
                BatteryLife.Visibility = Visibility.Visible;
                IsCharging.Visibility = Visibility.Hidden;
            }

            if (pw == "NoSystemBattery")
            {
                BatteryLife.Visibility = Visibility.Hidden;
                IsCharging.Visibility = Visibility.Hidden;
            }
        }
    }

}