using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Core;
using MyCompany.Interop.Extensibility;
using MyCompany.Interop.VBAExtensibility;

namespace MyVBEAddin
{
    [ComVisible(true), Guid("E70CC4E5-9A53-4402-B8DF-DB0FF6076035"), ProgId("MyVBAAddin.Connect")]
    public class Connect :IDTExtensibility2
    {
        private VBE _vbe;
        private AddIn _addin;
        private CommandBarButton _cbb;


        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _vbe = Application as VBE;
                _addin = AddInInst as AddIn;

                switch (ConnectMode)
                {
                    case ext_ConnectMode.ext_cm_AfterStartup:
                        InitializeAddIn();
                        break;
                    case ext_ConnectMode.ext_cm_Startup:
                        break;
                    case ext_ConnectMode.ext_cm_External:
                        break;
                    case ext_ConnectMode.ext_cm_CommandLine:
                        break;
                    default:
                        throw new ArgumentOutOfRangeException("ConnectMode");
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private IntPtr oldProc;
        private IntPtr chdHwnd;
        private void InitializeAddIn()
        {
            String STANDARD_COMMANDBAR_NAME = "Standard";
            String MENUBAR_COMMANDBAR_NAME = "Menu Bar";
            String TOOLS_COMMANDBAR_NAME = "Tools";
            String CODE_WINDOW_COMMANDBAR_NAME = "Code Window";

            try
            {
                CommandBar toolsCommandBar = _vbe.CommandBars[TOOLS_COMMANDBAR_NAME];
                
                _cbb = AddCommandBarButton(toolsCommandBar);

                _cbb.Click += cbb_Click;

                IntPtr hwnd;
                hwnd = NativeMethods.FindWindow("MsoCommandBarPopup", "Tools");
                hwnd = NativeMethods.GetWindow(hwnd, NativeMethods.GW_HWNDPREV);

                hwnd = NativeMethods.FindWindowEx(hwnd, IntPtr.Zero, "MDIClient", "");

                IntPtr chdHwnd = NativeMethods.FindWindowEx(hwnd, IntPtr.Zero, "VbaWindow", "WIAWrapper (程式碼)");

                oldProc = NativeMethods.SetWindowLongPtr64(new HandleRef(this, chdHwnd), NativeMethods.GWL_WNDPROC, new WindowProcEventHandler(WindowProc));
                
                Debug.WriteLine(_addin.ProgId + " loaded in VBA editor version " + _vbe.Version + " vbaWindow:" + hwnd.ToString());
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public IntPtr WindowProc(IntPtr hwnd, uint msg, IntPtr WParam, IntPtr LParam)
        {
            // WM_CHAR
            if (msg == 0x102)
            {
                int chr = WParam.ToInt32();

                if (chr == '\x9')
                {
                    //MessageBox.Show("Tab", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    int x, y, w, h;
                    _vbe.ActiveCodePane.GetSelection(out x, out y, out h, out w);

                    Debug.WriteLine(String.Format("x : {0}, y : {1}, h : {2}, w : {3}", x, y, h, w));
                    CodeModule cm = _vbe.ActiveCodePane.CodeModule;
                    string line = cm.Lines[x, 1];
                    Debug.WriteLine("the code :" + line);

                    line = line.TrimEnd();

                    string token = findTokenReverse(line);

                    if (token == "if")
                    {
                        // paste if stat
                        line = line.Substring(0, line.Length - 2);
                        String[] strs = new String[3];
                        var l = line.Length;

                        strs[0] = line + "If    Then";
                        strs[1] = "";
                        strs[2] = new String(' ', line.Length) + "End If";

                        cm.DeleteLines(x);
                        cm.InsertLines(x, String.Join("\r\n", strs));

                        // set pos
                        _vbe.ActiveCodePane.SetSelection(x, y, h, w);
                    }
                }

                Debug.WriteLine("the chr : " +  Convert.ToChar(chr));
            }

            if (msg == 0x14)
            {
                Debug.WriteLine("EraseBackgroud");
            }

            if (msg == 0x85)
            {
                Debug.WriteLine("NC PAINT");
            }
            return NativeMethods.CallWindowProc(oldProc, hwnd, msg, WParam, LParam);
        }

        private string findTokenReverse(string line)
        {
            string[] tokens = {"if"};

            Regex reg = new Regex(@"(\s\w+|^\w+)");

            var m = reg.Match(line);

            if (m.Success)
            {
                return m.Groups[m.Groups.Count - 1].Value.TrimStart().ToLower();
            }
            return String.Empty;
        }

        public delegate IntPtr WindowProcEventHandler(IntPtr hwnd, uint msg, IntPtr WParam, IntPtr LParam);

        void cbb_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            Debug.WriteLine("You Click me");
            Form frm = new Form();
            frm.Text = "My Custom VBA add-in";
            frm.SuspendLayout();

            Label lbl = new Label();
            lbl.Text = "In the codepane, you type [tab] and the token before [tab] is \"if\" then it will create the snippet for you.";
            lbl.Left = 10;
            lbl.Top = 10;
            frm.Controls.Add(lbl);

            frm.Show();
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            if (oldProc != IntPtr.Zero)
            {
                NativeMethods.SetWindowLongPtr64(new HandleRef(this, chdHwnd), NativeMethods.GWL_WNDPROC, oldProc);
            }

            switch (RemoveMode)
            {
                case ext_DisconnectMode.ext_dm_HostShutdown:
                    if (_cbb != null) _cbb.Delete();
                    //MessageBox.Show("ext_dm_HostShutdown", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
                case ext_DisconnectMode.ext_dm_UserClosed:
                    if (_cbb != null) _cbb.Delete();
                    //MessageBox.Show("ext_dm_UserClosed", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
                default:
                    throw new ArgumentOutOfRangeException("RemoveMode");
            }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
         
        }

        public void OnStartupComplete(ref Array custom)
        {
            InitializeAddIn();
        }

        public void OnBeginShutdown(ref Array custom)
        {
           
        }

        public CommandBarButton AddCommandBarButton(CommandBar cb)
        {
            CommandBarButton cbb;
            CommandBarControl cbc;

            cbc = cb.Controls.Add(MsoControlType.msoControlButton);
            cbb = cbc as CommandBarButton;

            cbb.Caption = "My Add-in Help";
            cbb.FaceId = 59;

            return cbb;
        }
    }
}
