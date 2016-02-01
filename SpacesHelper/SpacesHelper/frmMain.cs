using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using WindowsDesktop;
using WindowsInput;

namespace SpacesHelper {
    public partial class frmMain : Form {
        public frmMain() {
            InitializeComponent();
        }

        [DllImport("user32.dll")]
        public static extern int FindWindowEx(int hwndParent, int hwndChildAfter, string strClassName, string strWindowName);

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hwnd, ref Rect rectangle);

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);

        public const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        public const int MOUSEEVENTF_RIGHTUP = 0x10;

        private readonly KeyboardHook hook = new KeyboardHook();

        public struct Rect {
            public int Left { get; set; }
            public int Top { get; set; }
            public int Right { get; set; }
            public int Bottom { get; set; }

            public override string ToString() {
                return Left + " - " + Top + " - " + Bottom + " - " + Right;
            }
        }

        private string GetActiveWindowTitle() {
            const int nChars = 256;
            StringBuilder Buff = new StringBuilder(nChars);
            IntPtr handle = GetForegroundWindow();
            return GetWindowText(handle, Buff, nChars) > 0 ? Buff.ToString() : null;
        }


        private void MoveToDesktop(int target, bool moveWith) {

            new Thread(() => {
                Screen relevantScreen = Screen.FromHandle(GetForegroundWindow());

                while (GetActiveWindowTitle() != "Task View"){
                    InputSimulator.SimulateModifiedKeyStroke(VirtualKeyCode.LWIN, VirtualKeyCode.TAB);
                    Thread.Sleep(100);
                }

                while (!relevantScreen.Equals(Screen.FromHandle(GetForegroundWindow()))){
                    Cursor.Position = new Point(relevantScreen.WorkingArea.Top - 10, relevantScreen.WorkingArea.Right + 10);
                    var X = (uint) Cursor.Position.X;
                    var Y = (uint) Cursor.Position.Y;
                    mouse_event(MOUSEEVENTF_RIGHTDOWN | MOUSEEVENTF_RIGHTUP, X, Y, 0, 0);
                    Thread.Sleep(200);
                }

                Color lookFor = Color.FromArgb(43, 43, 43);
                var found = false;
                while (found == false){
                    InputSimulator.SimulateModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.F10);
                    Thread.Sleep(100);
                    var bmpScreenshot = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height, PixelFormat.Format32bppArgb);
                    var gfxScreenshot = Graphics.FromImage(bmpScreenshot);
                    gfxScreenshot.CopyFromScreen(Screen.PrimaryScreen.Bounds.X, Screen.PrimaryScreen.Bounds.Y, 0, 0, Screen.PrimaryScreen.Bounds.Size, CopyPixelOperation.SourceCopy);
                    found = checkImage(ResizeImage(bmpScreenshot, 480, 320), lookFor);
                }

                if (relevantScreen.Equals(Screen.FromHandle(GetForegroundWindow())) && GetActiveWindowTitle() == "Task View"){
                    InputSimulator.SimulateKeyPress(VirtualKeyCode.DOWN);
                    InputSimulator.SimulateKeyPress(VirtualKeyCode.DOWN);
                    InputSimulator.SimulateKeyPress(VirtualKeyCode.DOWN);
                    InputSimulator.SimulateKeyPress(VirtualKeyCode.RIGHT);
                    for (var i = 0; i < (GetCurrentDesktop() > target ? target - 1 : target - 2); i++){
                        InputSimulator.SimulateKeyPress(VirtualKeyCode.DOWN);
                        Thread.Sleep(50);
                    }
                    InputSimulator.SimulateKeyPress(VirtualKeyCode.RETURN);
                    Thread.Sleep(10);
                    InputSimulator.SimulateKeyPress(VirtualKeyCode.ESCAPE);
                    Thread.Sleep(10);
                    if (moveWith){
                        Thread.Sleep(100);
                        InputSimulator.SimulateModifiedKeyStroke(new[] {VirtualKeyCode.CONTROL, VirtualKeyCode.LWIN}, (GetCurrentDesktop() > target ? VirtualKeyCode.LEFT : VirtualKeyCode.RIGHT));
                    }
                }
            }).Start();
        }

        private bool checkImage(Bitmap b, Color col) {
            var c = GetSection(b, new Rectangle(0, 0, b.Width, b.Height));
            for (var x = 0; x < c.GetUpperBound(0); x++) {
                for (var y = 0; y < c.GetUpperBound(1); y++) {
                    if (col == c[x, y]){
                        return true;
                    }

                }
            }
            return false;
        }


        Color[,] GetSection(Image img, Rectangle rect) {
            var r = new Color[rect.Width, rect.Height];
            using (Bitmap b = new Bitmap(img)) { 
                BitmapData bd = b.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb); 
                var arr = new int[b.Width * b.Height - 1]; 
                Marshal.Copy(bd.Scan0, arr, 0, arr.Length);
                b.UnlockBits(bd); 
                for (var i = 0; i < arr.Length; i++) {
                    r[i % rect.Width, i / rect.Width] = Color.FromArgb(arr[i]); 
                }
            }
            return r; 
        }


        private int GetCurrentDesktop() {
            var i = 0;
            foreach (var virtualDesktop in VirtualDesktop.GetDesktops()){
                i++;
                if (virtualDesktop == VirtualDesktop.Current){
                    return i;
                }

            }
            return -1;
        }


        private void Form1_Load(object sender, EventArgs e) {
            hook.KeyPressed += Hook_KeyPressed;
            hook.RegisterHotKey(SpacesHelper.ModifierKeys.Alt, Keys.C);
            hook.RegisterHotKey(SpacesHelper.ModifierKeys.Alt, Keys.X);
        }


        void Hook_KeyPressed(object sender, KeyPressedEventArgs e) {    
            if (e.Key.ToString() == "C"){
                MoveToDesktop(GetCurrentDesktop() + 1, true);
            } else {
                MoveToDesktop(GetCurrentDesktop() - 1, true);
            }
        }


        public static Bitmap ResizeImage(Image image, int width, int height) {
            var destRect = new Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);
            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);
            using (var graphics = Graphics.FromImage(destImage)) {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                using (var wrapMode = new ImageAttributes()) {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }
            return destImage;
        }
    }
}
