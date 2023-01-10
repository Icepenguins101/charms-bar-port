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
using System.Timers;

namespace CharmsBarPort
{
    public partial class CharmsClock : Window
    {
        public CharmsClock()
        {
            // Here we obtain the current primary screen resolution. 
            // Setting neccesary values to ensure the window is shown correctly.
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
            Left = 33;
            Top = dispHeight - 190;
            //Debug.WriteLine($"Starting Charms Bar with a height of: {Height}px, and a left location at: {Left}px ");
            BrushConverter converter = new();
            var brush = (Brush)converter.ConvertFromString("#dd111111");
            Background = brush;

            InitializeComponent();
        }

        protected override void OnActivated(EventArgs e)
        {
                var MyDate = DateTime.Today.ToString("MMMM d");
                var MyWeek = DateTime.Today.ToString("dddd");
                var MyTime = DateTime.Now.ToString("hh:mm");
                Date.Content = MyDate;
                Week.Content = MyWeek;
                Clock.Content = MyTime;
            base.OnActivated(e);
        }

        protected override void OnDeactivated(EventArgs e)
        {
            base.OnDeactivated(e);
        }
    }
}
