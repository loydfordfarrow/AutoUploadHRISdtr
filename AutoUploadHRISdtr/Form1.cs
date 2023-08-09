using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.IO;

using CefSharp;
using CefSharp.WinForms;
using System.Reflection;
using System.Runtime.InteropServices;

namespace AutoUploadHRISdtr
{
    public partial class Form1 : Form
    {
        public ChromiumWebBrowser browser;
        public string url = "https://hris.api.com.ph/";
        //public string url = "http://192.168.10.253/";
        public string email = "loydfordfarrow@gmail.com";
        public string password = "4270540479";
        
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]

        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, IntPtr dwExtraInfo);
        const int MOUSEEVENTF_LEFTDOWN = 0x00000002;
        const int MOUSEEVENTF_LEFTUP = 0x00000004;
        const int MOUSEEVENTF_WHEEL = 0x0800;

        public int screenWidth = Screen.PrimaryScreen.Bounds.Width;
        public int screenHeight = Screen.PrimaryScreen.Bounds.Height;

        public void DoMouseClick()
        {
            int X = System.Windows.Forms.Cursor.Position.X;
            int Y = System.Windows.Forms.Cursor.Position.Y;
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, X, Y, 0, (IntPtr)0);
        }
        public void MouseWheelScroll()
        {
            int X = System.Windows.Forms.Cursor.Position.X;
            int Y = System.Windows.Forms.Cursor.Position.Y;
            mouse_event(MOUSEEVENTF_WHEEL, X, Y, -160, (IntPtr)0);
        }

        public int count1 = 5
            , count2 = 4
            , count3 = 4
            , count4 = 4
            , count5 = 4
            , count6 = 4
            , count7 = 4
            , count8 = 4
        ;


        public Form1()
        {
            InitializeComponent();

            //MessageBox.Show("screenWidth: " + screenWidth, "Sreen Resolution", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            // OPEN WEBSITE 
            Cef.Initialize(new CefSettings());
            browser = new ChromiumWebBrowser(url);
            this.Controls.Add(browser);
            browser.Dock = DockStyle.Fill;
            browser.LoadingStateChanged += (sender, args) =>
            {
                if (args.IsLoading == false)
                {
                    if (screenWidth == 1280 || screenWidth == 1024) { browser.SetZoomLevel(-1); }
                    //browser.ShowDevTools();
                    browser.EvaluateScriptAsync("document.querySelector('input[id=email]').value='" + email + "';");
                    browser.EvaluateScriptAsync("document.querySelector('input[id=password]').value='" + password + "';");
                    browser.EvaluateScriptAsync("setTimeout(function(){  document.querySelector('button[id=sign_in]').click()  },1000);");
                    // GO TO ATTENDANCE 
                    timer1 = new Timer();
                    timer1.Tick += new EventHandler(timer1_Tick);
                    timer1.Interval = 1000;
                    timer1.Start();
                }
            };
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            count1--;
            if (count1 == 0)
            {
                timer1.Stop();

                browser.EvaluateScriptAsync("setTimeout(function(){  document.querySelector('a[href=attendance]').click()  },1000);");
                // OPEN UPLOAD MODAL 
                timer1 = new Timer();
                timer1.Tick += new EventHandler(timer2_Tick);
                timer1.Interval = 1000;
                timer1.Start();
            }
        }
        private void timer2_Tick(object sender, EventArgs e)
        {
            count2--;
            if (count2 == 0)
            {
                timer1.Stop();

                browser.EvaluateScriptAsync("setTimeout(function(){  Array.from(document.querySelectorAll('a')).forEach(function(val){  if(val.innerText=='Upload'){val.style.color='red'; val.click();}  });  },2000);");
                // 
                timer1 = new Timer();
                timer1.Tick += new EventHandler(timer3_Tick);
                timer1.Interval = 1000;
                timer1.Start();
            }
        }
        private void timer3_Tick(object sender, EventArgs e)
        {
            count3--;
            if (count3 == 0)
            {
                timer1.Stop(); //browser.EvaluateScriptAsync("setTimeout(function(){  Array.from(document.querySelectorAll('span[id]')).forEach(function(e){  e.setAttribute('title','AGL_xxxx.txt'); e.innerText='AGL_xxxx.txt';  });  },1000);");
                
                // FILE TYPE 
                if (screenWidth == 1366) { Cursor.Position = new Point(520, 160); }
                if (screenWidth >= 1536) { Cursor.Position = new Point(760, 200); }
                DoMouseClick();
                // 
                timer1 = new Timer();
                timer1.Tick += new EventHandler(timer4_Tick);
                timer1.Interval = 1000;
                timer1.Start();
            }
        }
        private void timer4_Tick(object sender, EventArgs e)
        {
            count4--;
            if (count4 == 0)
            {
                timer1.Stop();

                // MOVE MOUSE POINT TO SCROLL 
                if (screenWidth == 1366) { Cursor.Position = new Point(520, 415); }
                if (screenWidth == 1536) { Cursor.Position = new Point(760, 350); }
                MouseWheelScroll();
                // 
                timer1 = new Timer();
                timer1.Tick += new EventHandler(timer5_Tick);
                timer1.Interval = 1000;
                timer1.Start();
            }
        }
        private void timer5_Tick(object sender, EventArgs e)
        {
            count5--;
            if (count5 == 0)
            {
                timer1.Stop();

                // SELECT FILE TYPE 
                if (screenWidth == 1366) { Cursor.Position = new Point(520, 415); }
                if (screenWidth == 1536) { Cursor.Position = new Point(760, 520); }
                DoMouseClick();
                // CHOOSE FILE 
                if (screenWidth == 1366) { Cursor.Position = new Point(520, 220); }
                if (screenWidth == 1536) { Cursor.Position = new Point(760, 280); }
                DoMouseClick();
                // 
                timer1 = new Timer();
                timer1.Tick += new EventHandler(timer6_Tick);
                timer1.Interval = 1000;
                timer1.Start();
            }
        }
        private void timer6_Tick(object sender, EventArgs e)
        {
            count6--;
            if (count6 == 0)
            {
                timer1.Stop();

                // SELECT FILE 
                if (screenWidth == 1366) { Cursor.Position = new Point(270, 160); }
                if (screenWidth == 1536) { Cursor.Position = new Point(335, 250); }
                DoMouseClick();
                DoMouseClick();
                // 
                timer1 = new Timer();
                timer1.Tick += new EventHandler(timer7_Tick);
                timer1.Interval = 1000;
                timer1.Start();
            }
        }
        private void timer7_Tick(object sender, EventArgs e)
        {
            count7--;
            if (count7 == 0)
            {
                timer1.Stop();

                // CLICK UPLOAD BUTTON 
                if (screenWidth == 1366) { Cursor.Position = new Point(520, 285); }
                if (screenWidth == 1536) { Cursor.Position = new Point(760, 350); }
                DoMouseClick();
                // 
                timer1 = new Timer();
                timer1.Tick += new EventHandler(timer8_Tick);
                timer1.Interval = 1000;
                timer1.Start();
            }
        }
        private void timer8_Tick(object sender, EventArgs e)
        {
            count8--;
            if (count8 == 0)
            {
                timer1.Stop();

                // CLICK UPLOAD BUTTON 
                if (screenWidth == 1280 || screenWidth == 1024) { Cursor.Position = new Point(520, 285); }
                if (screenWidth == 1366) { Cursor.Position = new Point(520, 285); }
                if (screenWidth == 1536) { Cursor.Position = new Point(760, 350); }
                DoMouseClick();
                // MOVE MOUSE POINTER 
                Cursor.Position = new Point(0, 0);
                // EXIT 
                Application.Exit();
            }
        }
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            //
            count1 = 20;
            count2 = 4;
            count3 = 4;
            count4 = 4;
            count5 = 4;
            count6 = 4;
            count7 = 4;
            count8 = 4;
            //
            timer1 = new Timer();
            timer1.Tick += new EventHandler(timer1_Tick);
            timer1.Interval = 1000;
            timer1.Start();
        }
    }
}
