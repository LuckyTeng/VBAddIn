using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace MyVBEAddin
{
    public class NativeMethods
    {
        /* GetNextWindow Constant */
        public const uint  
            GW_HWNDFIRST = 0,
            GW_HWNDLAST = 1,
            GW_HWNDNEXT = 2,
            GW_HWNDPREV = 3,
            GW_OWNER = 4,
            GW_CHILD = 5,
            GW_MAX = 5;

        public const int GWL_WNDPROC = -4;

        [DllImport("user32.dll", EntryPoint="FindWindowEx", CharSet=CharSet.Auto)]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll", EntryPoint="FindWindowEx", CharSet=CharSet.Auto)]
        public static extern IntPtr FindWindowEx( uint hwndParent, uint hwndChildAfter, string lpszClass, string lpszWindow);

        // This helper static method is required because the 32-bit version of user32.dll does not contain this API
        // (on any versions of Windows), so linking the method will fail at run-time. The bridge dispatches the request
        // to the correct function (GetWindowLong in 32-bit mode and GetWindowLongPtr in 64-bit mode)
        [DllImport("user32.dll", EntryPoint="SetWindowLong")]
        public static extern int SetWindowLong32(HandleRef hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll", EntryPoint="SetWindowLongPtr")]
        public static extern IntPtr SetWindowLongPtr64(HandleRef hWnd, int nIndex, [MarshalAs(UnmanagedType.FunctionPtr)] Delegate dwNewLong);

        [DllImport("user32.dll", EntryPoint="SetWindowLongPtr")]
        public static extern IntPtr SetWindowLongPtr64(HandleRef hWnd, int nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll", CharSet=CharSet.Ansi)]
        public static extern IntPtr FindWindow(string className, string windowName);

        [DllImport("user32.dll")]
        public static extern IntPtr CallWindowProc(IntPtr lpPrevWndFunc, IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);


        [DllImport("user32.dll", EntryPoint="GetWindow")]
        public static extern IntPtr GetWindow(IntPtr hWnd, uint wCmd);


    }
}
