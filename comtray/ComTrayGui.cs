using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.InteropServices;
using System.Drawing;


namespace comtray
{
    /// <summary>
    /// Summary description for NotifyIconForm.
    /// </summary>
    /// 
    public class ComTrayGui : System.Windows.Forms.Form
    {
        private System.Windows.Forms.NotifyIcon notifyIcon1;

        //forms vrs threading timer
        private System.Windows.Forms.Timer DebounceTimer;
        int sn_width;

        private System.ComponentModel.IContainer components;

        private ContextMenuStrip myRightClickMenu;

        private CodedRichTextBox richTextBox;

        private List<ComportScanner.DeviceInfo> curList;

        private void calculateDelta()
        {
            ComportScanner.DeviceInfo newPort;
            ComportScanner.DeviceInfo oldPort;
            int idxNew;
            int idxOld;
            List<ComportScanner.DeviceInfo> newList;
            List<ComportScanner.DeviceInfo> oldList;

            // get updated list
            newList = ComportScanner.GetComportList();

            // Sort of "merge" the two, old and new.
            // producing the current list
            //   Things that where removed = Code -1 // red-strikeout
            //   Things that where added   = Code +1 // green
            //   Things that are the same  = Code  0
            oldList = this.curList;
            this.curList = new List<ComportScanner.DeviceInfo>();

            idxNew = 0;
            idxOld = 0;

            while ((idxNew < newList.Count) && (idxOld < oldList.Count))
            {
                newPort = newList[idxNew];
                oldPort = oldList[idxOld];
                if (oldPort.for_application_use < 0)
                {
                    /* this was previously removed and remains removed */
                    /* so skip it */
                    idxOld++;
                    continue;
                }

                int r;
                r = ComportScanner.CompareDeviceInfo(newPort, oldPort);
                if (r == 0)
                {
                    newPort.for_application_use = 0;
                    this.curList.Add(newPort);
                    idxNew++;
                    idxOld++;
                    continue;
                }
                if (r < 0)
                {
                    /* NEW goes before, so consider this ADDED */
                    newPort.for_application_use = +1;
                    this.curList.Add(newPort);
                    idxNew++;
                    continue;
                }
                /* R > 0, then B must have been removed */
                oldPort.for_application_use = -1;
                this.curList.Add(oldPort);
                idxOld++;
                continue;
            }

            /* if we have more new items */
            while (idxNew < newList.Count)
            {
                ComportScanner.DeviceInfo tmp;
                tmp = newList[idxNew];
                tmp.for_application_use = +1;
                this.curList.Add(tmp);
                idxNew++;
            }

            /* and any that might have left */
            while( idxOld < oldList.Count)
            {
                ComportScanner.DeviceInfo tmp;
                tmp = oldList[idxOld];
                if (tmp.for_application_use < 0)
                {
                    /* it was already removed */
                }
                else
                {
                    tmp.for_application_use = -1;
                    this.curList.Add(tmp);
                }
                idxOld++;
            }
            /* done */
        }



        private void Usb_DeviceRemoved()
        {
            this.richTextBox.Text = "Debounce ...";
            this.DebounceTimer.Stop();
            this.DebounceTimer.Start();
        }

        private void Usb_DeviceAdded()
        {
            this.richTextBox.Text = "Debounce ...";
            this.DebounceTimer.Stop();
            this.DebounceTimer.Start();
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            if (m.Msg == UsbNotification.WmDevicechange)
            {
                switch ((int)m.WParam)
                {
                    case UsbNotification.DbtDeviceremovecomplete:
                        Usb_DeviceRemoved(); // this is where you do your magic
                        break;
                    case UsbNotification.DbtDevicearrival:
                        Usb_DeviceAdded(); // this is where you do your magic
                        break;
                }
            }
        }

        public ComTrayGui()
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();

            this.Icon = global::comtray.Properties.Resources.comtray_PLX_icon;
            //
            // TODO: Add any constructor code after InitializeComponent call
            //
            this.notifyIcon1.DoubleClick += new System.EventHandler(this.notifyIcon1_Click);
            this.Resize += new System.EventHandler(this.NotifyIconForm_Resize);
            this.FormClosing += NotifyIconForm_FormClosing;

            this.curList = ComportScanner.GetComportList();
            calculateDelta();
            updateTextAny();
            
            UsbNotification.RegisterUsbDeviceNotification(this.Handle);
        }

        static char get_ch(ComportScanner.DeviceInfo pThis)
        {
            if( pThis.for_application_use == 0)
            {
                return ' ';
            }
            if(pThis.for_application_use > 0)
            {
                return '+';
            }
            return '-';

        }

        private string _get_fmt( char ch, string a, string b, string c )
        {
            string s;
            s = String.Format("{0} {1} | {2} | {3}\r\n", ch,a.PadRight(6),b.PadRight(this.sn_width), c);
            return s;
        }

        private string get_str(char c, ComportScanner.DeviceInfo pThis)
        {
            return _get_fmt(c,pThis.name, pThis.usb_info.serialnumber, pThis.description);
        }

        private string get_sep()
        {
            return _get_fmt(' ',"Name", "SerialNumber", "Description");
        }

        private void updateTextAny()
        {
            updateTextBox(this.richTextBox);
        }
        private void updateTextBox(CodedRichTextBox tb )
        {
            char c;
            float zoomFactor;

            /* maintain the zoom factor */
            zoomFactor = tb.ZoomFactor;


            tb._Paint = false;
            tb.Text = "";
            Color cRed = Color.FromName("red");
            Color thisColor = cRed;
            
            int n;
            /* title */
            tb.CodedText(get_sep(), 0);
            /* body */

            /* somethings have HUGE serial numbers */
            this.sn_width = 12;
            foreach (ComportScanner.DeviceInfo tmp in this.curList)
            {
                if( this.sn_width < tmp.usb_info.serialnumber.Length)
                {
                    this.sn_width = tmp.usb_info.serialnumber.Length;
                }
            }

            n = 0;
            foreach (ComportScanner.DeviceInfo tmp in this.curList)
            {
                n++;
                if ((n % 16) == 0)
                {
                    /* sometimes I do have lots of serial ports
                     * happens when testing lots of IOT devices
                     */
                    tb.CodedText(get_sep(), 0);
                }
                c = get_ch(tmp);
                tb.CodedText(get_str(c,tmp), tmp.for_application_use);
            }
            // bug work around, set zoom factor to 0 first.
            tb.ZoomFactor = 1.0F;
            tb.ZoomFactor = zoomFactor;
            tb._Paint = true;
        }

        private void NotifyIconForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            show_form(false);
        }

        private void NotifyIconForm_Resize(object sender, EventArgs e)
        {
            int action = -1;
            if (WindowState == FormWindowState.Minimized)
            {
                action = 0;
            }
            else if (WindowState == FormWindowState.Maximized)
            {
                action = -1;
            }
            else if (WindowState == FormWindowState.Normal)
            {
                action = 1;
            }
            if (action == 1)
            {
                show_form(true);
            }
            else if (action == 0)
            {
                show_form(false);
            }

        }



        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        private void DebounceTick( object sender, EventArgs e )
        {
            // Some devices like dev boards have multiple UARTS
            // ie: JLINK on board debuggers have 3.
            // ie: Ti-Launch pads have 2
            // ie: FTDI chips either have 1 or 2 or 4
            // We want to treat the arrival as ONE arrivial
            // so we DEBOUNCE them, then redraw
            this.DebounceTimer.Stop();
            calculateDelta();
            updateTextAny();
            popup_form();
        }


        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.richTextBox = new comtray.CodedRichTextBox();
            this.DebounceTimer = new System.Windows.Forms.Timer(this.components);
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.myRightClickMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.SuspendLayout();
            // 
            // richTextBox
            // 
            this.richTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richTextBox.Location = new System.Drawing.Point(0, 0);
            this.richTextBox.Name = "richTextBox";
            this.richTextBox.ReadOnly = true;
            this.richTextBox.Size = new System.Drawing.Size(292, 266);
            this.richTextBox.TabIndex = 1;
            this.richTextBox.Text = "";
            this.richTextBox.WordWrap = false;
            // 
            // DebounceTimer
            // 
            this.DebounceTimer.Interval = 1500;
            this.DebounceTimer.Tick += new System.EventHandler(this.DebounceTick);
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.ContextMenuStrip = this.myRightClickMenu;
            this.notifyIcon1.Icon = global::comtray.Properties.Resources.comtray_PLX_icon;
            this.notifyIcon1.Text = "Duane\'s ComTray";
            // 
            // myRightClickMenu
            // 
            this.myRightClickMenu.Name = "myRightClickMenu";
            this.myRightClickMenu.Size = new System.Drawing.Size(61, 4);
            // 
            // ComTrayGui
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(6, 16);
            this.ClientSize = new System.Drawing.Size(292, 266);
            this.Controls.Add(this.richTextBox);
            this.Name = "ComTrayGui";
            this.Text = "Duane\'s ComTray (C) 2021 Duane Ellis";
            this.ResumeLayout(false);

        }

        private void myContextMenu_OnShowEventHandler(object sender, EventArgs e)
        {
            show_form(true);
        }

        private void myContextMenu_OnRefresh( object sender, EventArgs e )
        {
            calculateDelta();
            show_form(true);
        }

        private void myContextMenu_OnAboutHandler(object sender, EventArgs e)
        {
            string message;
            message = "Duanes ComTray V1\r\n(C) 2021 Duane Ellis";
            string title;
            title = "About ComTray";
            MessageBox.Show(message, title);
        }

        private void myContextMenu_OnExitHandler(object sender, EventArgs e)
        {
            Application.Exit();
        }


        private async void popup_form()
        {
            if( this.Visible)
            {
                return;
            }
            show_form(true);
            await Task.Delay(2500);
            show_form(false);
        }


        public void show_form(Boolean b)
        {
            if(b)
            {
                updateTextAny();
            }
            notifyIcon1.Visible = !b;
            this.Visible = b;
            if (b)
            {
                this.WindowState = FormWindowState.Normal;
            }
        }


        private void notifyIcon1_Click(object sender, System.EventArgs e)
        {
            show_form(true);

        }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>

    }
}

