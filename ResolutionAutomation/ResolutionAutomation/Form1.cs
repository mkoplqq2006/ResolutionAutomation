using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace ResolutionAutomation
{
    public partial class Form1 : Form
    {
        private bool run = true;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;

            string scriptPath = Application.StartupPath + "/script";
            string screenshotsPath = Application.StartupPath + "/screenshots";
            if(!string.IsNullOrEmpty(textBox1.Text))
            {
                screenshotsPath += "/"+ textBox1.Text;
            }

            if (!Directory.Exists(scriptPath))
            {
                Directory.CreateDirectory(scriptPath);
            }
            if (!Directory.Exists(screenshotsPath))
            {
                Directory.CreateDirectory(screenshotsPath);
            }

            DirectoryInfo scriptFolder = new DirectoryInfo(scriptPath);
            this.progressBar1.Value = 0;
            this.progressBar1.Maximum = scriptFolder.GetFiles().Length;

            foreach (FileInfo file in scriptFolder.GetFiles())
            {
                this.progressBar1.Value += 1;
                this.label2.Text = this.progressBar1.Value + "/" + this.progressBar1.Maximum;
                Application.DoEvents();

                string name = file.Name.Replace(".py", "");
                string[] nameArr = name.Split('×');

                DEVMODE DevM = new DEVMODE();
                DevM.dmSize = (short)Marshal.SizeOf(typeof(DEVMODE));
                bool mybool;
                mybool = EnumDisplaySettings(null, 0, ref DevM);
                DevM.dmPelsWidth = int.Parse(nameArr[0]);
                DevM.dmPelsHeight = int.Parse(nameArr[1]);
                long result = ChangeDisplaySettings(ref DevM, 0);

                SetCursorPos(600, 0);

                Thread.Sleep((int)numericUpDown1.Value * 1000);

                if (name == "1024×768") {
                    Thread.Sleep((int)numericUpDown1.Value * 1000);
                }

                Image baseImage = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
                Graphics g = Graphics.FromImage(baseImage);
                g.CopyFromScreen(new Point(0, 0), new Point(0, 0), Screen.AllScreens[0].Bounds.Size);
                g.Dispose();
                baseImage.Save(screenshotsPath + "/" + file.Name.Replace(".py", ".png"), ImageFormat.Png);
            }

            if (checkBox1.Checked)
            {
                if (!string.IsNullOrEmpty(textBox1.Text))
                {
                    Process proc2 = new Process();
                    proc2.StartInfo.WorkingDirectory = string.Format(Application.StartupPath + "\\screenshots\\");
                    proc2.StartInfo.FileName = "screenshot2.py";
                    proc2.Start();
                    proc2.WaitForExit();
                } 
                else
                {
                    Process proc = new Process();
                    proc.StartInfo.WorkingDirectory = string.Format(Application.StartupPath + "\\screenshots\\");
                    proc.StartInfo.FileName = "screenshot.py";
                    proc.Start();
                    proc.WaitForExit();
                }
            }
            this.WindowState = FormWindowState.Normal;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct DEVMODE
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string dmDeviceName;
            public short dmSpecVersion;
            public short dmDriverVersion;
            public short dmSize;
            public short dmDriverExtra;
            public int dmFields;
            public short dmOrientation;
            public short dmPaperSize;
            public short dmPaperLength;
            public short dmPaperWidth;
            public short dmScale;
            public short dmCopies;
            public short dmDefaultSource;
            public short dmPrintQuality;
            public short dmColor;
            public short dmDuplex;
            public short dmYResolution;
            public short dmTTOption;
            public short dmCollate;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string dmFormName;
            public short dmLogPixels;
            public int dmBitsPerPel;
            public int dmPelsWidth;
            public int dmPelsHeight;
            public int dmDisplayFlags;
            public int dmDisplayFrequency;
        }
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern int ChangeDisplaySettings([In] ref DEVMODE lpDevMode, int dwFlags);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern bool EnumDisplaySettings(string lpszDeviceName, Int32 iModeNum, ref DEVMODE lpDevMode);
        [DllImport("user32.dll")]
        static extern int SetCursorPos(int x, int y);

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                System.Environment.Exit(0);
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            System.Environment.Exit(0);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string screenshotsPath = Application.StartupPath + "/screenshots";
            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                screenshotsPath += "/" + textBox1.Text;
            }
            if (!Directory.Exists(screenshotsPath))
            {
                Directory.CreateDirectory(screenshotsPath);
            }
            string name = this.comboBox1.Text;
            string[] nameArr = name.Split('×');

            DEVMODE DevM = new DEVMODE();
            DevM.dmSize = (short)Marshal.SizeOf(typeof(DEVMODE));
            bool mybool;
            mybool = EnumDisplaySettings(null, 0, ref DevM);
            DevM.dmPelsWidth = int.Parse(nameArr[0]);
            DevM.dmPelsHeight = int.Parse(nameArr[1]);
            long result = ChangeDisplaySettings(ref DevM, 0);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string scriptPath = Application.StartupPath + "/script";
            DirectoryInfo scriptFolder = new DirectoryInfo(scriptPath);
            foreach (FileInfo file in scriptFolder.GetFiles())
            {
                comboBox1.Items.Add(file.Name.Replace(".py", ""));
                comboBox1.SelectedIndex = 0;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string screenshotsPath = Application.StartupPath + "\\screenshots";
            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                screenshotsPath += "\\" + textBox1.Text;
            }
            System.Diagnostics.Process.Start("Explorer.exe", screenshotsPath);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
            Thread.Sleep(500);

            string screenshotsPath = Application.StartupPath + "/screenshots";
            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                screenshotsPath += "/" + textBox1.Text;
            }
            if (!Directory.Exists(screenshotsPath))
            {
                Directory.CreateDirectory(screenshotsPath);
            }
            string name = this.comboBox1.Text;

            Image baseImage = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            Graphics g = Graphics.FromImage(baseImage);
            g.CopyFromScreen(new Point(0, 0), new Point(0, 0), Screen.AllScreens[0].Bounds.Size);
            g.Dispose();
            baseImage.Save(screenshotsPath + "/" + name + ".png", ImageFormat.Png);

            this.progressBar1.Value = 1;
            this.progressBar1.Maximum = 1;
            this.label2.Text = this.progressBar1.Value + "/" + this.progressBar1.Maximum;

            if (checkBox1.Checked)
            {
                if (!string.IsNullOrEmpty(textBox1.Text))
                {
                    Process proc2 = new Process();
                    proc2.StartInfo.WorkingDirectory = string.Format(Application.StartupPath + "\\screenshots\\");
                    proc2.StartInfo.FileName = "screenshot2.py";
                    proc2.Start();
                    proc2.WaitForExit();
                }
                else
                {
                    Process proc = new Process();
                    proc.StartInfo.WorkingDirectory = string.Format(Application.StartupPath + "\\screenshots\\");
                    proc.StartInfo.FileName = "screenshot.py";
                    proc.Start();
                    proc.WaitForExit();
                }
            }
            if (checkBox2.Checked && comboBox1.SelectedIndex + 1 == comboBox1.Items.Count)
            {
                button4_Click(this, e);
            }
            if (checkBox2.Checked && comboBox1.SelectedIndex + 1 < comboBox1.Items.Count)
            {
                comboBox1.SelectedIndex++;
                button3_Click(this, e);

                SetCursorPos(600, 0);
            }
            this.WindowState = FormWindowState.Normal;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            Process pro = new Process();
            pro.StartInfo.FileName = "iexplore.exe";
            pro.StartInfo.Arguments = "https://github.com/mkoplqq2006/ResolutionAutomation/issues";
            pro.Start();
        }
    }
}
