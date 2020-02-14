using com.valloon.ValloonShot.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace com.valloon.ValloonShot
{
    public partial class Form1 : Form
    {

        private String extension = @"png";
        private readonly String path = Application.StartupPath + @"\";
        private readonly String path2 = Application.StartupPath + @"\_shot\";
        private Boolean StealthMode = true;
        private int interval1 = 10;
        private int interval2 = 1799;
        private int hotkey = (int)Keys.Space;
        private String[] exceptWindowTitles = null;

        public Form1()
        {
            InitializeComponent();
        }

        public static String CorrectFileNames(String input)
        {
            String filename = input.Replace('/', ' ').Replace('\\', ' ').Replace(':', '.').Replace('*', ' ').Replace('?', ' ').Replace('\"', ' ').Replace('<', ' ').Replace('>', ' ').Replace('|', ' ');
            if (filename.Length > 200) filename = filename.Substring(0, 200) + "...";

            //string illegal = "\"M\"\\a/ry/ h**ad:>> a\\/:*?\"| li*tt|le|| la\"mb.?";
            string regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
            Regex r = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
            filename = r.Replace(filename, "");

            return filename;
        }

        [DllImport("User32")]
        public static extern bool RegisterHotKey(
            IntPtr hWnd,
            int id,
            int fsModifiers,
            int vk
        );
        [DllImport("User32")]
        public static extern bool UnregisterHotKey(
            IntPtr hWnd,
            int id
        );

        public const int MOD_WIN = 0x8;
        public const int MOD_SHIFT = 0x4;
        public const int MOD_CONTROL = 0x2;
        public const int MOD_ALT = 0x1;
        public const int WM_HOTKEY = 0x312;

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

        private Boolean ExceptThis(String SubActiveWindowTitle)
        {
            if (SubActiveWindowTitle == null) return true;
            foreach (String one in exceptWindowTitles)
            {
                if (SubActiveWindowTitle.IndexOf(one, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }
            return false;
        }

        private void ScreenShot(String path, String filename)
        {
            DirectoryInfo dir = new DirectoryInfo(path);
            if (!dir.Exists) dir.Create();
            if (!path.EndsWith(@"\")) path += @"\";
            filename = path + filename;
            Size size = SystemInformation.PrimaryMonitorSize;
            Bitmap bitmap = new Bitmap(size.Width, size.Height);
            Graphics graphics = Graphics.FromImage(bitmap);
            graphics.CopyFromScreen(0, 0, 0, 0, size);
            graphics.Flush();
            bitmap.Save(filename, ImageFormat.Png);
            bitmap.Dispose();
            graphics.Dispose();
        }

        private void hotkeyShot()
        {
            DateTime time = DateTime.Now;
            String filename = time.ToString("yyyy-MM-dd  HH.mm.ss");
            try
            {
                String activeWindowTitle = GetActiveWindowTitle();
                if (activeWindowTitle != null)
                {
                    filename += "  " + CorrectFileNames(activeWindowTitle) + "." + extension;
                }
                else
                {
                    filename += "." + extension;
                }
                ScreenShot(path2, filename);
                notifyIcon1.ShowBalloonTip(0, "Success !\r\n", filename, ToolTipIcon.Info);
            }
            catch (Exception ex)
            {
                notifyIcon1.ShowBalloonTip(0, "Save failed !\r\n" + filename, ex.Message, ToolTipIcon.Error);
            }
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY && m.WParam == (IntPtr)0)
            {
                hotkeyShot();
            }
            base.WndProc(ref m);
            this.Hide();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Hide();
            this.Visible = false;
            if (!RegisterHotKey(this.Handle, 0, MOD_WIN + MOD_ALT, hotkey) && !StealthMode)
                MessageBox.Show("Set hotkey failed.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            startToolStripMenuItem_Click(sender, e);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            UnregisterHotKey(this.Handle, 0);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            DateTime time = DateTime.Now;
            String filename = time.ToString("yyyy-MM-dd  HH.mm.ss");
            try
            {
                String activeWindowTitle = GetActiveWindowTitle();
                if (ExceptThis(activeWindowTitle)) return;
                if (activeWindowTitle != null)
                {
                    filename += "  " + CorrectFileNames(activeWindowTitle) + "." + extension;
                }
                else
                {
                    filename += "." + extension;
                }
                ScreenShot(path + time.ToString("yyyy-MM-dd"), filename);
            }
            catch (Exception ex)
            {
                notifyIcon1.ShowBalloonTip(0, "Save failed !\r\n" + filename, ex.Message, ToolTipIcon.Error);
                throw ex;
            }

        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            DateTime time = DateTime.Now;
            String filename = time.ToString("yyyy-MM-dd  HH.mm.ss");
            try
            {
                String activeWindowTitle = GetActiveWindowTitle();
                if (ExceptThis(activeWindowTitle)) return;
                if (activeWindowTitle != null)
                {
                    filename += "  " + CorrectFileNames(activeWindowTitle) + "." + extension;
                }
                else
                {
                    filename += "." + extension;
                }
                ScreenShot(path + time.ToString("yyyy"), filename);
            }
            catch (Exception ex)
            {
                notifyIcon1.ShowBalloonTip(0, "Save failed !\r\n" + filename, ex.Message, ToolTipIcon.Error);
            }
        }

        private void startToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                String filename = path + "config.cfg";
                FileInfo fileInfo = new FileInfo(filename);
                if (fileInfo.Exists)
                {
                    StreamReader reader = File.OpenText(filename);
                    String line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (line.StartsWith("//")) continue;
                        String[] var = line.Split(new Char[] { '=' }, 2);
                        if (var[0] == "interval1")
                        {
                            interval1 = Convert.ToInt16(var[1]);
                        }
                        else if (var[0] == "interval2")
                        {
                            interval2 = Convert.ToInt16(var[1]);
                        }
                        else if (var[0] == "extension")
                        {
                            extension = var[1].Trim();
                        }
                        else if (var[0] == "expect")
                        {
                            exceptWindowTitles = var[1].Trim().Split(new String[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                        }
                        else if (var[0] == "stealth")
                        {
                            StealthMode = var[1].Trim() == "1";
                            this.notifyIcon1.Visible = !StealthMode;
                        }
                    }
                    reader.Close();
                    reader.Dispose();
                }
                else
                {
                    StreamWriter writer = new StreamWriter("config.cfg");
                    writer.Write(
@"interval1=10
interval2=1199
extension=png
expect=Program Manager");
                    writer.Close();
                }
            }
            catch (Exception ex)
            {
                if (!StealthMode)
                    MessageBox.Show("Read config failed. " + ex.Message, "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            notifyIcon1.ShowBalloonTip(0, "I am alive !", "Interval1=" + interval1 + ", Interval2=" + interval2, ToolTipIcon.Info);
            timer1.Interval = 1000 * interval1;
            timer2.Interval = 1000 * interval2;
            timer1.Enabled = true;
            timer2.Enabled = true;
            notifyIcon1.Icon = Resources.valloon;
            startToolStripMenuItem.Enabled = false;
            stopToolStripMenuItem.Enabled = true;
        }

        private void stopToolStripMenuItem_Click(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            //timer2.Enabled=false;
            notifyIcon1.Icon = Resources.valloon_gray;
            startToolStripMenuItem.Enabled = true;
            stopToolStripMenuItem.Enabled = false;
        }

        private void openConfigToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start("Notepad.exe", "config.cfg");
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UnregisterHotKey(this.Handle, 0);
            this.Close();
        }

        private void notifyIcon1_DoubleClick(object sender, EventArgs e)
        {
            if (startToolStripMenuItem.Enabled)
                startToolStripMenuItem_Click(sender, e);
            else
                stopToolStripMenuItem_Click(sender, e);
        }

        private void openTargetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe", @"file:///" + path);
        }
    }
}
