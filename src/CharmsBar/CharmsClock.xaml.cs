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
using static System.Resources.ResXFileRef;

namespace CharmsBarPort
{
    public partial class CharmsClock : Window
    {
        public Microsoft.Win32.RegistryKey localKey = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, RegistryView.Registry64);
        public int isAirPlaneOn = 0;

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
            Width = 456; // 80 to 90 is magic spot.
            WindowStartupLocation = WindowStartupLocation.Manual;
            Left = 51;
            Top = dispHeight - 188;
            BrushConverter converter = new();
            var brush = (Brush)converter.ConvertFromString("#f0111111");
            Background = brush;
            InitializeComponent();
            _initTimer();
        }

    private Timer t = null;
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
                Date.Content = DateTime.Today.ToString("MMMM d");
                Week.Content = DateTime.Today.ToString("dddd");
                Clocks.Content = DateTime.Now.ToString("h                           mm");
                Clocked.Content = DateTime.Now.ToString("mm");

                var dispWidth = SystemParameters.PrimaryScreenWidth;
                var dispHeight = SystemParameters.PrimaryScreenHeight;

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

                CheckBatteryStatus();
                var localKey = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, RegistryView.Registry64);
                var localKey2 = localKey.OpenSubKey("SYSTEM\\ControlSet001\\Control\\RadioManagement\\SystemRadioState");
                var isAirPlaneOn = localKey2.GetValue("", "").ToString();
                var nw = IsConnected().ToString();
                var nw2 = IsLocal().ToString();
                var nw3 = IsWeak().ToString();
                
                if (nw == "True" && nw2 == "False" && nw3 == "False" && isAirPlaneOn == "0")
                {
                    NoDrivers.Visibility = Visibility.Hidden;
                    NoInternetFound.Visibility = Visibility.Hidden;
                    Ethernet.Visibility = Visibility.Hidden;
                    HasInternet.Visibility = Visibility.Visible;
                    WeakInternet.Visibility = Visibility.Hidden;
                    Airplane.Visibility = Visibility.Hidden;
                }

                if (nw == "True" && nw2 == "False" && nw3 == "False" && isAirPlaneOn == "1")
                {
                    NoDrivers.Visibility = Visibility.Hidden;
                    NoInternetFound.Visibility = Visibility.Hidden;
                    Ethernet.Visibility = Visibility.Hidden;
                    HasInternet.Visibility = Visibility.Hidden;
                    WeakInternet.Visibility = Visibility.Hidden;
                    Airplane.Visibility = Visibility.Visible;
                }

                if (nw2 == "True" && nw == "False" && nw3 == "False" && isAirPlaneOn == "0")
                {
                    NoDrivers.Visibility = Visibility.Hidden;
                    NoInternetFound.Visibility = Visibility.Hidden;
                    Ethernet.Visibility = Visibility.Visible;
                    HasInternet.Visibility = Visibility.Hidden;
                    WeakInternet.Visibility = Visibility.Hidden;
                    Airplane.Visibility = Visibility.Hidden;
                }

                if (nw == "False" && nw2 == "False" && nw3 == "False" && isAirPlaneOn == "0")
                {
                    NoDrivers.Visibility = Visibility.Hidden;
                    NoInternetFound.Visibility = Visibility.Visible;
                    Ethernet.Visibility = Visibility.Hidden;
                    HasInternet.Visibility = Visibility.Hidden;
                    WeakInternet.Visibility = Visibility.Hidden;
                    Airplane.Visibility = Visibility.Hidden;
                }

                if (nw == "False" && nw2 == "False" && nw3 == "True" && isAirPlaneOn == "0")
                {
                    NoDrivers.Visibility = Visibility.Hidden;
                    NoInternetFound.Visibility = Visibility.Hidden;
                    Ethernet.Visibility = Visibility.Hidden;
                    HasInternet.Visibility = Visibility.Hidden;
                    WeakInternet.Visibility = Visibility.Visible;
                    Airplane.Visibility = Visibility.Hidden;
                }

                if (nw == "False" && nw2 == "False" && nw3 == "False" && isAirPlaneOn == "1")
                {
                    NoDrivers.Visibility = Visibility.Hidden;
                    NoInternetFound.Visibility = Visibility.Hidden;
                    Ethernet.Visibility = Visibility.Hidden;
                    HasInternet.Visibility = Visibility.Hidden;
                    WeakInternet.Visibility = Visibility.Hidden;
                    Airplane.Visibility = Visibility.Visible;
                }

                if (Clocks.Content.ToString().Length < 30 && Clocks.Content.ToString().StartsWith("1 ") == false && Clocks.Content.ToString().StartsWith("10") == false && Clocks.Content.ToString().StartsWith("12") == false || Clocks.Content.ToString().Length == 30 && Clocks.Content.ToString().StartsWith("1 ") == false && Clocks.Content.ToString().StartsWith("10") == false && Clocks.Content.ToString().StartsWith("12") == false)
                {
                    Clocks.Margin = new Thickness(94, 3, 0, -106);
                    ClockLines.Margin = new Thickness(138, -24.99, -190, -98);
                    Clocked.Margin = new Thickness(157, -17, -190, -198);
                    Week.Margin = new Thickness(267, 2, 0, -18);
                    Date.Margin = new Thickness(267, 4, 0, -24);
                }

                if (Clocks.Content.ToString().Length < 30 && Clocks.Content.ToString().StartsWith("1 ") == true || Clocks.Content.ToString().Length == 30 && Clocks.Content.ToString().StartsWith("1 ") == true)
                {
                    Clocks.Margin = new Thickness(95, 3, 0, -106);
                    ClockLines.Margin = new Thickness(125, -24.99, -190, -98);
                    Clocked.Margin = new Thickness(144, -17, -190, -198);
                    Week.Margin = new Thickness(255, 2, 0, -18);
                    Date.Margin = new Thickness(255, 4, 0, -24);
                }

                if (Clocks.Content.ToString().Length == 31 && Clocks.Content.ToString().StartsWith("10") == true || Clocks.Content.ToString().Length == 30 && Clocks.Content.ToString().StartsWith("10") == true)
                {
                    Clocks.Margin = new Thickness(95, 3, 0, -106);
                    ClockLines.Margin = new Thickness(169, -24.99, -190, -98);
                    Clocked.Margin = new Thickness(188, -17, -190, -198);
                    Week.Margin = new Thickness(298, 2, 0, -18);
                    Date.Margin = new Thickness(300, 4, 0, -24);
                    this.Width = 450 + Date.Content.ToString().Length + 30;
                }

                if (Clocks.Content.ToString().Length == 31 && Clocks.Content.ToString().StartsWith("12") == true || Clocks.Content.ToString().Length == 30 && Clocks.Content.ToString().StartsWith("12") == true)
                {
                    Clocks.Margin = new Thickness(95, 3, 0, -106);
                    ClockLines.Margin = new Thickness(169, -24.99, -190, -98);
                    Clocked.Margin = new Thickness(188, -17, -190, -198);
                    Week.Margin = new Thickness(298, 2, 0, -18);
                    Date.Margin = new Thickness(300, 4, 0, -24);
                    this.Width = 450 + Date.Content.ToString().Length + 30;
                }

                if (Clocks.Content.ToString().Length == 31 && Clocks.Content.ToString().StartsWith("10") == false && Clocks.Content.ToString().StartsWith("12") == false)
                {
                    Clocks.Margin = new Thickness(94, 3, 0, -106);
                    ClockLines.Margin = new Thickness(157, -24.99, -190, -98);
                    Clocked.Margin = new Thickness(175, -17, -190, -198);
                    Week.Margin = new Thickness(287, 2, 0, -18);
                    Date.Margin = new Thickness(287, 4, 0, -24);
                }

                if (Date.Content.ToString().Length > 11)
                {
                    if (Clocks.Content.ToString().Length == 30)
                    {
                        this.Width = 450 + Date.Content.ToString().Length + 36;
                    }

                    if (Clocks.Content.ToString().Length == 31 && Clocks.Content.ToString().StartsWith("10") == false && Clocks.Content.ToString().StartsWith("12") == false)
                    {
                        this.Width = 450 + Date.Content.ToString().Length + 46;
                    }
                }

                if (Date.Content.ToString().Length < 11)
                {
                    if (Clocks.Content.ToString().Length == 30)
                    {
                        this.Width = 450 + Date.Content.ToString().Length;
                    }

                    if (Clocks.Content.ToString().Length == 31 && Clocks.Content.ToString().StartsWith("10") == false && Clocks.Content.ToString().StartsWith("12") == false)
                    {
                        this.Width = 450 + Date.Content.ToString().Length + 26;
                    }
                }

                ClockBorder.Width = this.Width;
                localKey.Close();
            }));
        }
        private bool IsConnected()
        {
            var connectionProfile = NetworkInformation.GetInternetConnectionProfile();
            return (connectionProfile != null &&
                  (connectionProfile.GetNetworkConnectivityLevel() == NetworkConnectivityLevel.InternetAccess));
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

        private void CheckBatteryStatus()
        {
            var pw = SystemInformation.PowerStatus.BatteryChargeStatus.ToString();
            var pw2 = SystemInformation.PowerStatus.PowerLineStatus.ToString();
            double pw3 = SystemInformation.PowerStatus.BatteryLifePercent;
            var dasBoot = pw3.ToString();

            if (dasBoot.StartsWith("1.0") == true || dasBoot.StartsWith("0.9") == true && pw3 > 0.96 && dasBoot.StartsWith("1.0") == false)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/BatteryFull.png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.9") == true && pw3 > 0.95 && pw3 < 0.96 && dasBoot.StartsWith("1.0") == false)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery90.png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.9") == true && pw3 < 0.95)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery90.png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.8") == true && pw3 > 0.86)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery80.png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.8") == true && pw3 < 0.86)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery80.png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.7") == true && pw3 > 0.76)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery70.png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.7") == true && pw3 < 0.76)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery70.png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.6") == true && pw3 > 0.66)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery60.png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.6") == true && pw3 < 0.66)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery60.png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.5") == true && pw3 > 0.56)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery50.png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.5") == true && pw3 < 0.56)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery50.png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.4") == true && pw3 > 0.46)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery40.png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.4") == true && pw3 < 0.46)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery40.png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.3") == true && pw3 > 0.36)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery30.png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.3") == true && pw3 < 0.36)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery30.png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.2") == true && pw3 > 0.26)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery20.png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.2") == true && pw3 < 0.26)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery20.png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.1") == true && pw3 > 0.16)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery10.png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.1") == true && pw3 < 0.16)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery1.png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.1") == true && pw3 > 0.6)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery5.png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.1") == true && pw3 < 0.6)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery5.png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.1") == true && pw3 > 0.2)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery1.png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.1") == true && pw3 < 0.2)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery1.png", UriKind.Relative));
            }

            if (dasBoot.StartsWith("0.1") == true && pw3 < 0.1)
            {
                BatteryLife.Source = new BitmapImage(new Uri(@"/Assets/Images/Battery0.png", UriKind.Relative));
            }

            if (pw2 == "Online")
            {
                IsCharging.Visibility = Visibility.Visible;
            }

            if (pw2 == "Offline")
            {
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