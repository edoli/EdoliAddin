using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    public class KeyboardHook
    {
        public delegate int LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private static LowLevelKeyboardProc _proc = HookCallback;
        private static IntPtr _hookID = IntPtr.Zero;

        private static long timestamp = 0;

        private static bool isCalled;
        private const int WH_KEYBOARD = 2;
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private const int HC_ACTION = 0;

        private static System.IO.StreamWriter logFile = new System.IO.StreamWriter("C:/Users/daniel/Downloads/test/log.txt");


        #region VirtualKey
        public enum VKeys : int
        {
            VK_LBUTTON = 0x01,  //Left mouse button
            VK_RBUTTON = 0x02,  //Right mouse button
            VK_CANCEL = 0x03,  //Control-break processing
            VK_MBUTTON = 0x04,  //Middle mouse button (three-button mouse)
            VK_BACK = 0x08,  //BACKSPACE key
            VK_TAB = 0x09,  //TAB key
            VK_CLEAR = 0x0C,  //CLEAR key
            VK_RETURN = 0x0D,  //ENTER key
            VK_SHIFT = 0x10,  //SHIFT key
            VK_CONTROL = 0x11,  //CTRL key
            VK_MENU = 0x12,  //ALT key
            VK_PAUSE = 0x13,  //PAUSE key
            VK_CAPITAL = 0x14,  //CAPS LOCK key
            VK_HANGUL = 0x15,
            VK_ESCAPE = 0x1B,  //ESC key
            VK_SPACE = 0x20,  //SPACEBAR
            VK_PRIOR = 0x21,  //PAGE UP key
            VK_NEXT = 0x22,  //PAGE DOWN key
            VK_END = 0x23,  //END key
            VK_HOME = 0x24,  //HOME key
            VK_LEFT = 0x25,  //LEFT ARROW key
            VK_UP = 0x26,  //UP ARROW key
            VK_RIGHT = 0x27,  //RIGHT ARROW key
            VK_DOWN = 0x28,  //DOWN ARROW key
            VK_SELECT = 0x29,  //SELECT key
            VK_PRINT = 0x2A,  //PRINT key
            VK_EXECUTE = 0x2B,  //EXECUTE key
            VK_SNAPSHOT = 0x2C,  //PRINT SCREEN key
            VK_INSERT = 0x2D,  //INS key
            VK_DELETE = 0x2E,  //DEL key
            VK_HELP = 0x2F,  //HELP key
            VK_0 = 0x30,  //0 key
            VK_1 = 0x31,  //1 key
            VK_2 = 0x32,  //2 key
            VK_3 = 0x33,  //3 key
            VK_4 = 0x34,  //4 key
            VK_5 = 0x35,  //5 key
            VK_6 = 0x36,  //6 key
            VK_7 = 0x37,  //7 key
            VK_8 = 0x38,  //8 key
            VK_9 = 0x39,  //9 key
            VK_A = 0x41,  //A key
            VK_B = 0x42,  //B key
            VK_C = 0x43,  //C key
            VK_D = 0x44,  //D key
            VK_E = 0x45,  //E key
            VK_F = 0x46,  //F key
            VK_G = 0x47,  //G key
            VK_H = 0x48,  //H key
            VK_I = 0x49,  //I key
            VK_J = 0x4A,  //J key
            VK_K = 0x4B,  //K key
            VK_L = 0x4C,  //L key
            VK_M = 0x4D,  //M key
            VK_N = 0x4E,  //N key
            VK_O = 0x4F,  //O key
            VK_P = 0x50,  //P key
            VK_Q = 0x51,  //Q key
            VK_R = 0x52,  //R key
            VK_S = 0x53,  //S key
            VK_T = 0x54,  //T key
            VK_U = 0x55,  //U key
            VK_V = 0x56,  //V key
            VK_W = 0x57,  //W key
            VK_X = 0x58,  //X key
            VK_Y = 0x59,  //Y key
            VK_Z = 0x5A,  //Z key
            VK_NUMPAD0 = 0x60,  //Numeric keypad 0 key
            VK_NUMPAD1 = 0x61,  //Numeric keypad 1 key
            VK_NUMPAD2 = 0x62,  //Numeric keypad 2 key
            VK_NUMPAD3 = 0x63,  //Numeric keypad 3 key
            VK_NUMPAD4 = 0x64,  //Numeric keypad 4 key
            VK_NUMPAD5 = 0x65,  //Numeric keypad 5 key
            VK_NUMPAD6 = 0x66,  //Numeric keypad 6 key
            VK_NUMPAD7 = 0x67,  //Numeric keypad 7 key
            VK_NUMPAD8 = 0x68,  //Numeric keypad 8 key
            VK_NUMPAD9 = 0x69,  //Numeric keypad 9 key
            VK_SEPARATOR = 0x6C,  //Separator key
            VK_SUBTRACT = 0x6D,  //Subtract key
            VK_DECIMAL = 0x6E,  //Decimal key
            VK_DIVIDE = 0x6F,  //Divide key
            VK_F1 = 0x70,  //F1 key
            VK_F2 = 0x71,  //F2 key
            VK_F3 = 0x72,  //F3 key
            VK_F4 = 0x73,  //F4 key
            VK_F5 = 0x74,  //F5 key
            VK_F6 = 0x75,  //F6 key
            VK_F7 = 0x76,  //F7 key
            VK_F8 = 0x77,  //F8 key
            VK_F9 = 0x78,  //F9 key
            VK_F10 = 0x79,  //F10 key
            VK_F11 = 0x7A,  //F11 key
            VK_F12 = 0x7B,  //F12 key
            VK_SCROLL = 0x91,  //SCROLL LOCK key
            VK_LSHIFT = 0xA0,  //Left SHIFT key
            VK_RSHIFT = 0xA1,  //Right SHIFT key
            VK_LCONTROL = 0xA2,  //Left CONTROL key
            VK_RCONTROL = 0xA3,  //Right CONTROL key
            VK_LMENU = 0xA4,   //Left MENU key
            VK_RMENU = 0xA5,  //Right MENU key
            VK_PLAY = 0xFA,  //Play key
            VK_ZOOM = 0xFB, //Zoom key
        }
        #endregion
     
        public static void SetHook()
        {
            _hookID = SetWindowsHookEx(WH_KEYBOARD, _proc, IntPtr.Zero, (uint)AppDomain.GetCurrentThreadId());
        }

        public static void ReleaseHook()
        {
            UnhookWindowsHookEx(_hookID);
        }

        private static int HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            //long newTimeStamp = DateTimeOffset.Now.ToUnixTimeMilliseconds();
            //bool throttle = newTimeStamp - timestamp < 10;

            //if (!throttle)
            //{
            //    timestamp = newTimeStamp;
            //}

            if (nCode < 0)
            {
                return (int)CallNextHookEx(_hookID, nCode, wParam, lParam);
            }
            else
            {
                VKeys key = (VKeys)wParam;

                short t1 = GetKeyState((int)VKeys.VK_CONTROL);
                short t2 = GetKeyState((int)VKeys.VK_MENU);
                short t3 = GetKeyState((int)VKeys.VK_SHIFT);

                bool ctrl = GetKeyState((int) VKeys.VK_CONTROL) < 0;
                bool alt = GetKeyState((int)VKeys.VK_MENU) < 0;
                bool shift = GetKeyState((int)VKeys.VK_SHIFT) < 0;

                // align
                if (ctrl && alt) {
                    switch (key)
                    {
                        case VKeys.VK_NUMPAD4:
                            AlignTool.AlignLeft();
                            break;
                        case VKeys.VK_NUMPAD6:
                            AlignTool.AlignRight();
                            break;
                        case VKeys.VK_NUMPAD8:
                            AlignTool.AlignTop();
                            break;
                        case VKeys.VK_NUMPAD2:
                            AlignTool.AlignBottom();
                            break;
                        case VKeys.VK_NUMPAD5:
                            AlignTool.AlignCenter();
                            break;
                        case VKeys.VK_NUMPAD7:
                            AlignTool.AlignInRow();
                            break;
                        case VKeys.VK_H:
                            AlignTool.AlignCenterHorizontal();
                            break;
                        case VKeys.VK_T:
                            AlignTool.AlignCenterVertical();
                            break;
                    }
                }

                //switch (key)
                //{
                //    case VKeys.VK_NUMPAD4:
                //        if (ctrl && alt) { AlignTool.AlignLeft(); }
                //        break;
                //    case VKeys.VK_G:
                //        AlignTool.AlignLeft();
                //        break;
                //}

                return (int)CallNextHookEx(_hookID, nCode, wParam, lParam);
            }
        }

        [DllImport("user32.dll")]
        private static extern bool keybd_event(uint bVk, uint bScan, uint dwFlags, int dwExtraInfo);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook,
            LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode,
            IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern short GetKeyState(int nCode);
    }
}
