/// <summary>
/// Provides wrapper around several API calls that manipulate application windows. This was built using snippets of code from various sources with custom
/// modifications with some new code feathered in. As a result, it's not very consistent or intuitive. For example, the API signatures are all over the place,
/// mixing different data types (int/uint/IntPtr, etc.), some with SetLastError, some with CharSet, some without.
///
/// In other words: This class needs a serious re-write!
/// </summary>
namespace ZoomMeetingBotSDK
{
    using System;
    using System.Drawing;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading;
    using System.Windows.Forms;
    using global::ZoomMeetingBotSDK.Interop.HostApp;
    using static Utils.ZMBUtils;


    /*
    public static class ProcessExtensions
    {
        public static IEnumerable<Process> GetChildProcesses(this Process process)
        {
            List<Process> children = new List<Process>();
            ManagementObjectSearcher mos = new ManagementObjectSearcher(string.Format("Select * From Win32_Process Where ParentProcessID={0}", process.Id));

            foreach (ManagementObject mo in mos.Get())
            {
                children.Add(Process.GetProcessById(Convert.ToInt32(mo["ProcessID"])));
            }

            return children;
        }
    }
    */

    internal class WindowTools
    {
        public enum ShowCmd : uint
        {
            SW_HIDE = 0,
            SW_MINIMIZE = 6,
            SW_SHOWNORMAL = 1,
            SW_SHOWMINIMIZED = 2,
            SW_SHOWMAXIMIZED = 3,
            SW_SHOWNOACTIVE = 4,
            SW_SHOW = 5,
            SW_SHOWMINNOACTIVE = 7,
            SW_SHOWNA = 8,
            SW_RESTORE = 9,
            SW_SHOWDEFAULT = 10,
            SW_FORCEMINIMIZE = 11,
        }

        public enum ScrollBarVal : uint
        {
            SB_LINEUP = 0,      // Scrolls one line up
            SB_LINELEFT = 0,    // Scrolls one cell left
            SB_LINEDOWN = 1,    // Scrolls one line down
            SB_LINERIGHT = 1,   // Scrolls one cell right
            SB_PAGEUP = 2,      // Scrolls one page up
            SB_PAGELEFT = 2,    // Scrolls one page left
            SB_PAGEDOWN = 3,    // Scrolls one page down
            SB_PAGERIGHT = 3,   // Scrolls one page right
            SB_PAGETOP = 6,     // Scrolls to the upper left
            SB_LEFT = 6,        // Scrolls to the left
            SB_PAGEBOTTOM = 7,  // Scrolls to the upper right
            SB_RIGHT = 7,       // Scrolls to the right
            SB_ENDSCROLL = 8,    // Ends scroll
        }

        private enum GetWindow_Cmd : uint
        {
            GW_HWNDFIRST = 0,
            GW_HWNDLAST = 1,
            GW_HWNDNEXT = 2,
            GW_HWNDPREV = 3,
            GW_OWNER = 4,
            GW_CHILD = 5,
            GW_ENABLEDPOPUP = 6,
        }

        private enum WinMsg : uint
        {
            WM_SYSCOMMAND = 0x0112,
            WM_CLOSE = 0x0010,
            WM_QUIT = 0x0012,
            WM_SCROLL = 0x0114,
            WM_VSCROLL = 0x0115,
        }

        private enum SysCmd : uint
        {
            SC_MAXIMIZE = 0xf030,
            SC_RESTORE = 0xf120,
        }

        [Flags]
        private enum MouseEventFlags : uint
        {
            MOUSEEVENTF_MOVE = 0x0001,
            MOUSEEVENTF_LEFTDOWN = 0x0002,
            MOUSEEVENTF_LEFTUP = 0x0004,
            MOUSEEVENTF_RIGHTDOWN = 0x0008,
            MOUSEEVENTF_RIGHTUP = 0x0010,
            MOUSEEVENTF_MIDDLEDOWN = 0x0020,
            MOUSEEVENTF_MIDDLEUP = 0x0040,
            MOUSEEVENTF_XDOWN = 0x0080,
            MOUSEEVENTF_XUP = 0x0100,
            MOUSEEVENTF_WHEEL = 0x0800,
            MOUSEEVENTF_VIRTUALDESK = 0x4000,
            MOUSEEVENTF_ABSOLUTE = 0x8000,
        }

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool ShowWindow(IntPtr hWnd, uint cmdShow);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr GetForegroundWindow();

        internal struct INPUT
        {
            public uint Type;
            public MOUSEKEYBDHARDWAREINPUT Data;
        }

        [StructLayout(LayoutKind.Explicit)]
        internal struct MOUSEKEYBDHARDWAREINPUT
        {
            [FieldOffset(0)]
            public MOUSEINPUT Mouse;
        }

        internal struct MOUSEINPUT
        {
            public int X;
            public int Y;
            public uint MouseData;
            public uint Flags;
            public uint Time;
            public IntPtr ExtraInfo;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, IntPtr lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(IntPtr lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder title, int size);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", EntryPoint = "SetWindowPos", SetLastError = true)]
        private static extern IntPtr SetWindowPos(IntPtr hWnd, int hWndInsertAfter, int x, int y, int cx, int cy, int wFlags);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool MoveWindow(IntPtr hWnd, int x, int y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll")]
        private static extern bool ClientToScreen(IntPtr hWnd, ref System.Drawing.Point lpPoint);

        [DllImport("user32.dll")]
        private static extern uint SendInput(uint nInputs, [MarshalAs(UnmanagedType.LPArray), In] INPUT[] pInputs, int cbSize);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(HandleRef hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;

            public RECT(int left, int top, int right, int bottom)
            {
                Left = left;
                Top = top;
                Right = right;
                Bottom = bottom;
            }

            public RECT(System.Drawing.Rectangle r)
                : this(r.Left, r.Top, r.Right, r.Bottom) { }

            public int X
            {
                get { return Left; }
                set { Right -= Left - value; Left = value; }
            }

            public int Y
            {
                get { return Top; }
                set { Bottom -= Top - value; Top = value; }
            }

            public int Height
            {
                get { return Bottom - Top; }
                set { Bottom = value + Top; }
            }

            public int Width
            {
                get { return Right - Left; }
                set { Right = value + Left; }
            }

            public System.Drawing.Point Location
            {
                get { return new System.Drawing.Point(Left, Top); }
                set { X = value.X; Y = value.Y; }
            }

            public System.Drawing.Size Size
            {
                get { return new System.Drawing.Size(Width, Height); }
                set { Width = value.Width; Height = value.Height; }
            }

            public static implicit operator System.Drawing.Rectangle(RECT r)
            {
                return new System.Drawing.Rectangle(r.Left, r.Top, r.Width, r.Height);
            }

            public static implicit operator RECT(System.Drawing.Rectangle r)
            {
                return new RECT(r);
            }

            public static bool operator ==(RECT r1, RECT r2)
            {
                return r1.Equals(r2);
            }

            public static bool operator !=(RECT r1, RECT r2)
            {
                return !r1.Equals(r2);
            }

            public bool Equals(RECT r)
            {
                return r.Left == Left && r.Top == Top && r.Right == Right && r.Bottom == Bottom;
            }

            public override bool Equals(object obj)
            {
                if (obj is RECT rect)
                {
                    return Equals(rect);
                }
                else if (obj is System.Drawing.Rectangle drawRect)
                {
                    return Equals(new RECT(drawRect));
                }

                return false;
            }

            public override int GetHashCode()
            {
                return ((System.Drawing.Rectangle)this).GetHashCode();
            }

            public override string ToString()
            {
                return string.Format(System.Globalization.CultureInfo.CurrentCulture, "{{Left={0},Top={1},Right={2},Bottom={3}}}", Left, Top, Right, Bottom);
            }
        }

#pragma warning disable 649

        public enum MouseButton
        {
            Left = 0,
            Middle = 1,
            Right = 3,
        }

#pragma warning restore 649
        /// <summary>
        /// This object controls locking of input we are sending to an external application.
        /// </summary>
        public static readonly object InputLock = new object();

        /// <summary>
        /// The previous position of the mouse cursor.  Used by ClickOnPoint() to restore the old position after performing the requested click.
        /// </summary>
        private static System.Drawing.Point oldPos;

        private bool isDisposed = false;

        public void Dispose() => Dispose(true);

        public virtual void Dispose(bool disposing)
        {
            if (isDisposed)
            {
                return;
            }

            if (disposing)
            {
                // TBD: Clean up whatever...
            }

            isDisposed = true;
        }

        public static void SavePos()
        {
            lock (InputLock)
            {
                oldPos = Cursor.Position;
            }
        }

        public static void RestorePos()
        {
            lock (InputLock)
            {
                Cursor.Position = oldPos;
            }
        }

        private static double Hypotenuse(double a, double b)
        {
            return Math.Sqrt((a * a) + (b * b));
        }

        /// <summary>
        /// Smoothy moves the mouse from current position to the given target position.
        /// </summary>
        private static void SmoothMoveMouse(System.Drawing.Point targetPos)
        {
            var moveDelay = 10;
            var origPos = Cursor.Position;
            double dx = Math.Abs(targetPos.X - origPos.X);
            double dy = Math.Abs(targetPos.Y - origPos.Y);
            //double h = Hypotenuse(dx, dy);
            //int moves = (int)(h / (Global.cfg.MouseMovementRate * pixPerTix));
            int moves = (int)Global.cfg.MouseMovementRate / moveDelay;
            dx = ((targetPos.X < origPos.X) ? -1 : 1) * (dx / moves);
            dy = ((targetPos.Y < origPos.Y) ? -1 : 1) * (dy / moves);

            double x = origPos.X;
            double y = origPos.Y;
            for (int i = 0; i < moves; i++)
            {
                Thread.Sleep(moveDelay);
                x += dx;
                y += dy;
                Cursor.Position = new System.Drawing.Point((int)x, (int)y);
                //Application.DoEvents();
            }
        }

        public static void WiggleMouse()
        {
            oldPos = Cursor.Position;
            Cursor.Position = new System.Drawing.Point(oldPos.X + 1, oldPos.Y + 0);
            Thread.Sleep(50);
            Cursor.Position = new System.Drawing.Point(oldPos.X + 2, oldPos.Y + 0);
            Thread.Sleep(50);
            Cursor.Position = new System.Drawing.Point(oldPos.X + 3, oldPos.Y + 0);
            Thread.Sleep(50);
            Cursor.Position = new System.Drawing.Point(oldPos.X + 4, oldPos.Y + 1);
            Thread.Sleep(50);
            Cursor.Position = new System.Drawing.Point(oldPos.X + 4, oldPos.Y + 2);
            Thread.Sleep(50);
            Cursor.Position = new System.Drawing.Point(oldPos.X + 4, oldPos.Y + 3);
            Thread.Sleep(50);
            Cursor.Position = new System.Drawing.Point(oldPos.X + 3, oldPos.Y + 4);
            Thread.Sleep(50);
            Cursor.Position = new System.Drawing.Point(oldPos.X + 2, oldPos.Y + 4);
            Thread.Sleep(50);
            Cursor.Position = new System.Drawing.Point(oldPos.X + 1, oldPos.Y + 4);
            Thread.Sleep(50);
            Cursor.Position = new System.Drawing.Point(oldPos.X + 0, oldPos.Y + 4);
            Thread.Sleep(50);
            Cursor.Position = new System.Drawing.Point(oldPos.X + 0, oldPos.Y + 3);
            Thread.Sleep(50);
            Cursor.Position = new System.Drawing.Point(oldPos.X + 0, oldPos.Y + 2);
            Thread.Sleep(50);
            Cursor.Position = new System.Drawing.Point(oldPos.X + 0, oldPos.Y + 1);
            Thread.Sleep(50);
            Cursor.Position = oldPos;
            Thread.Sleep(1350);
        }

        private static void Click(IntPtr wndHandle, System.Drawing.Point clientPoint, MouseButton btn = MouseButton.Left)
        {
            lock (InputLock)
            {
                MouseEventFlags nDnFlag;
                MouseEventFlags nUpFlag;
                switch (btn)
                {
                    case MouseButton.Left:
                        nDnFlag = MouseEventFlags.MOUSEEVENTF_LEFTDOWN;
                        nUpFlag = MouseEventFlags.MOUSEEVENTF_LEFTUP;
                        break;
                    case MouseButton.Middle:
                        nDnFlag = MouseEventFlags.MOUSEEVENTF_MIDDLEDOWN;
                        nUpFlag = MouseEventFlags.MOUSEEVENTF_MIDDLEUP;
                        break;
                    case MouseButton.Right:
                        nDnFlag = MouseEventFlags.MOUSEEVENTF_RIGHTDOWN;
                        nUpFlag = MouseEventFlags.MOUSEEVENTF_RIGHTUP;
                        break;
                    default:
                        throw new Exception(string.Format("Unsupported MouseButton Value: {0}", btn.ToString()));
                }

                // get screen coordinates
                //ClientToScreen(wndHandle, ref clientPoint);
                Global.hostApp.Log(LogType.DBG, "ClickOnPoint {0:X8} {1}", (uint)wndHandle, clientPoint.ToString());

                // set cursor on coords, and press mouse
                Cursor.Position = new System.Drawing.Point(clientPoint.X, clientPoint.Y);

                Thread.Sleep(Global.cfg.ClickDelayMilliseconds);

                var inputMouseDown = new INPUT
                {
                    Type = 0,
                };
                // input type mouse
                inputMouseDown.Data.Mouse.Flags = (uint)nDnFlag; // Button down

                var inputMouseUp = new INPUT
                {
                    Type = 0,
                };
                // input type mouse
                inputMouseUp.Data.Mouse.Flags = (uint)nUpFlag; // Button up

                var inputs = new INPUT[] { inputMouseDown, inputMouseUp };

                SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
            }
        }

        public static void ClickOnPoint(IntPtr wndHandle, System.Drawing.Point clientPoint, MouseButton btn = MouseButton.Left)
        {
            lock (InputLock)
            {
                //SavePos();
                //Click(wndHandle, clientPoint, btn);
                //RestorePos();
                WindowTools.ClientToScreen(wndHandle, ref clientPoint);
                SmoothMoveMouse(clientPoint);
                Click(wndHandle, clientPoint, btn);
            }
        }

        public static void ClickOnPoint(IntPtr wndHandle, System.Windows.Point clientPoint)
        {
            ClickOnPoint(wndHandle, new System.Drawing.Point((int)clientPoint.X, (int)clientPoint.Y));
        }

        public static void ClickOnPoint(int wndHandle, System.Windows.Point clientPoint)
        {
            ClickOnPoint((IntPtr)wndHandle, new System.Drawing.Point((int)clientPoint.X, (int)clientPoint.Y));
        }

        // Click in the middle of an object
        public static void ClickMiddle(IntPtr wndHandle, System.Windows.Rect clientRect)
        {
            var point = new System.Windows.Point(clientRect.X + (clientRect.Width / 2), clientRect.Y + (clientRect.Height / 2));
            WindowTools.ClickOnPoint(wndHandle, point);
        }

        public static void ClickMiddle(int wndHandle, System.Windows.Rect clientRect)
        {
            ClickMiddle((IntPtr)wndHandle, clientRect);
        }

        public static IntPtr FindWindowByClass(string lpClassName)
        {
            lock (InputLock)
            {
                return FindWindow(lpClassName, IntPtr.Zero);
            }
        }

        public static string GetWindowText(IntPtr hWnd, int size = 1024)
        {
            var sb = new StringBuilder(size);
            GetWindowText(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }

        public static IntPtr FindWindow(string lpClassName, Regex reWindowName, out string sWindowName)
        {
            lock (InputLock)
            {
                sWindowName = null;
                var h = FindWindow(lpClassName, IntPtr.Zero);
                while (h != IntPtr.Zero)
                {
                    var sbTitle = new StringBuilder(1024);
                    _ = GetWindowText(h, sbTitle, sbTitle.Capacity);
                    if (reWindowName.IsMatch(sbTitle.ToString()))
                    {
                        sWindowName = sbTitle.ToString();
                        break;
                    }
                    h = GetWindow(h, (uint)GetWindow_Cmd.GW_HWNDNEXT);
                }
                return h;
            }
        }

        public static IntPtr WaitWindow(string lpClassName, string lpWindowName, int timeout = 10000, int poll = 250)
        {
            lock (InputLock)
            {
                var h = IntPtr.Zero;
                var dtStart = DateTime.UtcNow;
                while (true)
                {
                    if (lpWindowName == null)
                    {
                        h = FindWindow(lpClassName, IntPtr.Zero);
                    }
                    else if (lpClassName == null)
                    {
                        h = FindWindow(IntPtr.Zero, lpWindowName);
                    }
                    else
                    {
                        h = FindWindow(lpClassName, lpWindowName);
                    }
                    if (h != IntPtr.Zero)
                    {
                        return h;
                    }
                    var dtNow = DateTime.UtcNow;
                    if (dtNow.Subtract(dtStart).TotalMilliseconds >= timeout)
                    {
                        throw new TimeoutException(string.Format("Timeout waiting for window with class {0} and name {1}", repr(lpClassName), repr(lpWindowName)));
                    }
                    Thread.Sleep(poll);
                }
            }
        }

        public static IntPtr WaitWindow(string lpClassName, Regex reWindowName, out string sWindowName, int timeout = 10000, int poll = 250)
        {
            lock (InputLock)
            {
                var h = IntPtr.Zero;
                var dtStart = DateTime.UtcNow;
                while (true)
                {
                    if (lpClassName == null)
                    {
                        h = FindWindow(string.Empty, reWindowName, out sWindowName);
                    }
                    else
                    {
                        h = FindWindow(lpClassName, reWindowName, out sWindowName);
                    }
                    if (h != IntPtr.Zero)
                    {
                        return h;
                    }
                    var dtNow = DateTime.UtcNow;
                    if (dtNow.Subtract(dtStart).TotalMilliseconds >= timeout)
                    {
                        throw new TimeoutException(string.Format("Timeout waiting for window with class {0} and name {1}", repr(lpClassName), repr(reWindowName.ToString())));
                    }
                    Thread.Sleep(poll);
                }
            }
        }

        public static void CloseWindow(IntPtr hWnd)
        {
            lock (InputLock)
            {
                SendMessage(hWnd, (uint)WinMsg.WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            }
        }

        public static void QuitWindow(IntPtr hWnd)
        {
            lock (InputLock)
            {
                SendMessage(hWnd, (uint)WinMsg.WM_QUIT, IntPtr.Zero, IntPtr.Zero);
            }
        }

        public static void RestoreWindow(IntPtr hWnd)
        {
            lock (InputLock)
            {
                SendMessage(hWnd, (uint)WinMsg.WM_SYSCOMMAND, (IntPtr)SysCmd.SC_RESTORE, IntPtr.Zero);
            }
        }

        /// <summary>
        /// Brings the provided window into focus if it is not already.
        /// </summary>
        /// <param name="hWnd">Handle of the window to bring into focus.</param>
        public static void FocusWindow(IntPtr hWnd)
        {
            // https://www.pinvoke.net/default.aspx/user32/SetForegroundWindow.html%20diff=y
            // https://shlomio.wordpress.com/2012/09/04/solved-setforegroundwindow-win32-api-not-always-works/
            if (hWnd == IntPtr.Zero)
            {
                return;
            }

            lock (InputLock)
            {
                if (GetForegroundWindow() != hWnd)
                {
                    SetForegroundWindow(hWnd);

                    // Give UI a chance to process it
                    Thread.Sleep(Global.cfg.UIActionDelayMilliseconds);
                }
            }
        }

        private static readonly string SENDKEYS_SPECIAL_CHARS = "+^%~()[]{}";

        public static void SendKeys(IntPtr hWnd, string keys, bool hide = false, bool literal = false)
        {
            // literal:
            //   The plus sign (+), caret (^), percent sign (%), tilde (~), and parentheses () have special meanings to SendKeys. To specify one of these characters, enclose it within braces ({}).
            //   For example, to specify the plus sign, use "{+}". To specify brace characters, use "{{}" and "{}}". Brackets ([ ]) have no special meaning to SendKeys, but you must enclose them
            //   in braces. In other applications, brackets do have a special meaning that might be significant when dynamic data exchange (DDE) occurs.
            if (literal)
            {
                StringBuilder sb = new StringBuilder();
                foreach (var ch in keys)
                {
                    if (ch == '\n')
                    {
                        // Multiple lines in Zoom are sent with SHIFT+ENTER
                        sb.Append("+{ENTER}");
                    }
                    else if (SENDKEYS_SPECIAL_CHARS.IndexOf(ch) == -1)
                    {
                        sb.Append(ch);
                    }
                    else
                    {
                        sb.Append("{" + ch + "}");
                    }
                }

                keys = sb.ToString();
            }

            // TBD: Move back to old app after we're done?
            lock (InputLock)
            {
                Global.hostApp.Log(LogType.DBG, "WindowTools.SendKeys {0}", repr(hide ? "(hidden)" : keys));
                FocusWindow(hWnd);
                System.Windows.Forms.SendKeys.SendWait(keys);

                // Give UI a chance to process it
                Thread.Sleep(Global.cfg.KeyboardInputDelayMilliseconds);
            }
        }

        public static void SendKeys(string keys)
        {
            SendKeys(IntPtr.Zero, keys);
        }

        /// <summary>
        /// Pastes the given text into the given window using the Windows Clipboard.  This is a more efficient method to send text than SendKeys() when
        /// the text to be set is more than a few characters.  Falls back on SendKeys if the clipboard operation fails.
        /// </summary>
        /// <param name="hWnd">Handle to the destination window.</param>
        /// <param name="text">Text to paste.</param>
        public static void SendText(IntPtr hWnd, string text)
        {
            lock (InputLock)
            {
                var useClipboard = !Global.cfg.DisableClipboardPasteText;

                Global.hostApp.Log(LogType.DBG, $"SendText useClipboard={repr(useClipboard)} text={repr(text)}");

                Exception caughtException = null;
                if (useClipboard)
                {
                    // The clipboard has to be accessed via an STA thread, so create a temporary one for this call
                    Thread thread = new Thread(() =>
                    {
                        try
                        {
                            Clipboard.SetText(text);
                        }
                        catch (Exception ex)
                        {
                            caughtException = ex;
                        }
                    });
                    thread.SetApartmentState(ApartmentState.STA);
                    thread.Start();
                    thread.Join();

                    if (caughtException == null)
                    {
                        FocusWindow(hWnd);
                        System.Windows.Forms.SendKeys.SendWait("^V");

                        // Give UI a chance to process it
                        Thread.Sleep(Global.cfg.KeyboardInputDelayMilliseconds);

                        return;
                    }
                    else
                    {
                        // Sometimes SetText() throws an error in KERNELBASE.dll.  Seems to happen when Chrome Remote Desktop is in use
                        //Global.hostApp.Log(LogType.ERR, "Caught exception while trying to set clipboard text: {0}", repr(caughtException.ToString()));
                        Global.hostApp.Log(LogType.WRN, "Failed to set clipboard text; Falling back on SendKeys");
                    }
                }

                SendKeys(hWnd, text, false, true);
            }
        }

        public static void SendText(string text)
        {
            SendText(IntPtr.Zero, text);
        }

        public static void SetWindowSize(IntPtr hWnd, Rectangle rect)
        {
            lock (InputLock)
            {
                MoveWindow(hWnd, (int)rect.X, (int)rect.Y, (int)rect.Width, (int)rect.Height, true);
            }
        }

        public static System.Windows.Rect GetWindowRect(IntPtr hWnd)
        {
            HandleRef h = new HandleRef(null, hWnd);
            _ = GetWindowRect(h, out RECT rect);
            return new System.Windows.Rect(rect.X, rect.Y, rect.Width, rect.Height);
        }

        public static void ScrollWindow(IntPtr hWnd, ScrollBarVal scrollBarVal)
        {
            SendMessage(hWnd, (uint)WinMsg.WM_SCROLL, (IntPtr)scrollBarVal, IntPtr.Zero);
            SendMessage(hWnd, (uint)WinMsg.WM_SCROLL, (IntPtr)ScrollBarVal.SB_ENDSCROLL, IntPtr.Zero);
        }

        /// <summary>
        /// Tries to wake up the screen and kill any running screen saver.
        /// </summary>
        public static void WakeScreen()
        {
            Global.hostApp.Log(LogType.DBG, "Trying to wake the screen");
            try
            {
                WiggleMouse();
            }
            catch (Exception ex)
            {
                Global.hostApp.Log(LogType.WRN, "Failed to wake screen: {0}", ex.ToString());
            }

            if (ScreenSaver.GetScreenSaverRunning())
            {
                Global.hostApp.Log(LogType.DBG, "Trying to stop the screen saver");
                try
                {
                    ScreenSaver.KillScreenSaver();
                }
                catch (Exception ex)
                {
                    Global.hostApp.Log(LogType.WRN, "Failed to stop screen saver: {0}", ex.ToString());
                }
            }
            else
            {
                Global.hostApp.Log(LogType.DBG, "Screen saver is not running");
            }
        }

        public static class ScreenSaver
        {
            // From here: https://www.codeproject.com/Articles/17067/Controlling-The-Screen-Saver-With-C

            // Signatures for unmanaged calls
            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            private static extern bool SystemParametersInfo(int uAction, int uParam, ref int lpvParam, int flags);

            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            private static extern bool SystemParametersInfo(int uAction, int uParam, ref bool lpvParam, int flags);

            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            private static extern int PostMessage(IntPtr hWnd, int wMsg, int wParam, int lParam);

            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            private static extern IntPtr OpenDesktop(string hDesktop, int flags, bool inherit, uint desiredAccess);

            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            private static extern bool CloseDesktop(IntPtr hDesktop);

            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            private static extern bool EnumDesktopWindows(IntPtr hDesktop, EnumDesktopWindowsProc callback, IntPtr lParam);

            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            private static extern bool IsWindowVisible(IntPtr hWnd);

            // Callbacks
            private delegate bool EnumDesktopWindowsProc(IntPtr hDesktop, IntPtr lParam);

            // Constants
            private const int SPI_GETSCREENSAVERACTIVE = 16;
            private const int SPI_SETSCREENSAVERACTIVE = 17;
            private const int SPI_GETSCREENSAVERTIMEOUT = 14;
            private const int SPI_SETSCREENSAVERTIMEOUT = 15;
            private const int SPI_GETSCREENSAVERRUNNING = 114;
            private const int SPIF_SENDWININICHANGE = 2;

            private const uint DESKTOP_WRITEOBJECTS = 0x0080;
            private const uint DESKTOP_READOBJECTS = 0x0001;
            private const int WM_CLOSE = 16;

            // Returns TRUE if the screen saver is active
            // (enabled, but not necessarily running).
            public static bool GetScreenSaverActive()
            {
                bool isActive = false;

                SystemParametersInfo(SPI_GETSCREENSAVERACTIVE, 0, ref isActive, 0);
                return isActive;
            }

            // Pass in TRUE(1) to activate or FALSE(0) to deactivate
            // the screen saver.
            public static void SetScreenSaverActive(int active)
            {
                int nullVar = 0;

                SystemParametersInfo(SPI_SETSCREENSAVERACTIVE, active, ref nullVar, SPIF_SENDWININICHANGE);
            }

            // Returns the screen saver timeout setting, in seconds
            public static int GetScreenSaverTimeout()
            {
                int value = 0;

                SystemParametersInfo(SPI_GETSCREENSAVERTIMEOUT, 0, ref value, 0);
                return value;
            }

            // Pass in the number of seconds to set the screen saver
            // timeout value.
            public static void SetScreenSaverTimeout(int value)
            {
                int nullVar = 0;

                SystemParametersInfo(SPI_SETSCREENSAVERTIMEOUT, value, ref nullVar, SPIF_SENDWININICHANGE);
            }

            // Returns TRUE if the screen saver is actually running
            public static bool GetScreenSaverRunning()
            {
                bool isRunning = false;

                SystemParametersInfo(SPI_GETSCREENSAVERRUNNING, 0, ref isRunning, 0);
                return isRunning;
            }

            // From Microsoft's Knowledge Base article #140723:
            // http://support.microsoft.com/kb/140723
            // "How to force a screen saver to close once started
            // in Windows NT, Windows 2000, and Windows Server 2003"

            public static void KillScreenSaver()
            {
                IntPtr hDesktop = OpenDesktop("Screen-saver", 0, false, DESKTOP_READOBJECTS | DESKTOP_WRITEOBJECTS);
                if (hDesktop != IntPtr.Zero)
                {
                    EnumDesktopWindows(hDesktop, new EnumDesktopWindowsProc(KillScreenSaverFunc), IntPtr.Zero);
                    CloseDesktop(hDesktop);
                }
                else
                {
                    PostMessage(GetForegroundWindow(), WM_CLOSE, 0, 0);
                }
            }

            private static bool KillScreenSaverFunc(IntPtr hWnd, IntPtr lParam)
            {
                if (IsWindowVisible(hWnd))
                {
                    PostMessage(hWnd, WM_CLOSE, 0, 0);
                }

                return true;
            }
        }
    }
}
