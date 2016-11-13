using BingWallpaper.Models;
using Microsoft.Win32;
using System;
using System.IO;
using System.Timers;
using System.Runtime.InteropServices;

using System.Windows.Forms;
using System.Drawing;

namespace BingWallpaper
{
    class Program : Form
    {
        #region unsafe
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;
        #endregion
        
        [STAThread]
        static void Main()
        {
            Application.Run(new Program());


            // Hide console
            var handle = GetConsoleWindow();
            ShowWindow(handle, SW_HIDE);

            Log.Print("bing wallpaper process started");
            IImageProvider provider = new BingDayImageProvider("en-CY");
            Log.Print("setting program to run on startup");
            SetStartup();

            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = 1000 * 60 * 60 * 24; // 1000ms/sec -> 60sec/min -> 60min/hr -> 24hr/day
            timer.AutoReset = true;
            timer.Enabled = true;
            timer.Elapsed += (s, e) => SetWallpaper(provider);
            timer.Start();

            // Set wallpaper on first run
            SetWallpaper(provider);

            // Keep process alive
            Console.Read();
        }

        private NotifyIcon trayIcon;
        private ContextMenu trayMenu;

        public Program()
        {
            // Create a simple tray menu with only one item.
            trayMenu = new ContextMenu();
            trayMenu.MenuItems.Add("Exit", OnExit);

            // Create a tray icon. In this example we use a
            // standard system icon for simplicity, but you
            // can of course use your own custom icon too.
            trayIcon = new NotifyIcon();
            trayIcon.Text = "BingWallpaper";
            trayIcon.Icon = new Icon(SystemIcons.Application, 40, 40);

            // Add menu to tray icon and show it.
            trayIcon.ContextMenu = trayMenu;
            trayIcon.Visible = true;
        }

        protected override void OnLoad(EventArgs e)
        {
            Visible = false; // Hide form window.
            ShowInTaskbar = false; // Remove from taskbar.

            base.OnLoad(e);
        }

        private void OnExit(object sender, EventArgs e)
        {
            Application.Exit();
        }

        protected override void Dispose(bool isDisposing)
        {
            if (isDisposing)
            {
                // Release the icon resource.
                trayIcon.Dispose();
            }

            base.Dispose(isDisposing);
        }

        static void SetWallpaper(IImageProvider provider)
        {
            if (provider == null)
                throw new ArgumentNullException(nameof(provider));
            try
            {
                Log.Print("asking for image uri");
                var uri = provider.Uri().Result;
                Log.Print($"fetching image from \"{uri.ToString()}\"");
                Wallpaper.Set(uri, Wallpaper.Style.Stretched);
                Log.Print("wallpaper has been updated");
            }
            catch (Exception e)
            {
                Log.Print("failed to updated wallpaper");
                Log.Print(e);
            }
        }

        static void SetStartup()
        {
            RegistryKey rk = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);

            rk.SetValue("BingWallpaper", Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "BingWallpaper.exe"));
        }
    }
}
