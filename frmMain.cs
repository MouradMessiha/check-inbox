using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace CheckInbox
{
    public partial class frmMain : Form
    {
        [DllImport("User32")]
        private static extern int SetForegroundWindow(IntPtr hwnd);     
        [DllImportAttribute("User32.DLL")]     
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool IsIconic(IntPtr hWnd);
        private const int SW_SHOWNOACTIVATE = 4;
        private const int SW_MINIMIZE = 6;

        private Thread mobjCheckNewEmailsThread;
        private Thread mobjUpdateNotifyIconThread;
        private bool mblnFormClosed = false;
        private bool mblnLoggedIn = false;
        private bool mblnNotifyIconOn = false;

        Microsoft.Office.Interop.Outlook.Application mobjOutlookApp;
        Microsoft.Office.Interop.Outlook._NameSpace mobjOutlookNamespace;
        Microsoft.Office.Interop.Outlook.Folder mobjOutlookInbox;

        List<string> mlstEntryIDsViewed;

        public frmMain()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close(); 
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            mlstEntryIDsViewed = new List<string>();
            objNotifyIcon.Icon = Properties.Resources.email_red;
            mobjCheckNewEmailsThread = new Thread(StartCheckNewEmailsThread);
            mobjCheckNewEmailsThread.Start();
            mobjUpdateNotifyIconThread = new Thread(StartUpdateNotifyIconThread);
            mobjUpdateNotifyIconThread.Start();
            Hide();
        }

        private void StartCheckNewEmailsThread()
        {
            while (!mblnFormClosed)
            {
                try 
                {
                    if (!mblnLoggedIn)
                    {
                        mobjOutlookApp = new Microsoft.Office.Interop.Outlook.Application();
                        mobjOutlookNamespace = mobjOutlookApp.GetNamespace("MAPI");
                        mobjOutlookNamespace.Logon(null, null, false, false);
                        mobjOutlookInbox = (Microsoft.Office.Interop.Outlook.Folder)mobjOutlookNamespace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
                        mblnLoggedIn = true;
                        mblnNotifyIconOn = false;
                    }
                    SetIconFlag();
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    mblnLoggedIn = false;
                }
                // sleep 
                for (int intCount = 0; intCount < 30; intCount++)  // sleep 3 seconds
                {
                    if (!mblnFormClosed)
                    {
                        Thread.Sleep(100);
                    }
                }
            }
        }

        private void SetIconFlag()
        {
            bool blnUnreadMail = false;
            foreach (Object objItem in mobjOutlookInbox.Items)
            {
                if (objItem is Microsoft.Office.Interop.Outlook.MailItem)
                {
                    if (((Microsoft.Office.Interop.Outlook.MailItem)objItem).UnRead)
                        blnUnreadMail = true;
                    else
                        if (!mlstEntryIDsViewed.Contains(((Microsoft.Office.Interop.Outlook.MailItem)objItem).EntryID)) // email was not there when icon was last clicked to open outlook, could be an email that was not read by user, but flagged as read because the inbox folder is open and email is highlighted
                            blnUnreadMail = true;
                }
                else if (objItem is Microsoft.Office.Interop.Outlook.AppointmentItem)
                {
                    if (((Microsoft.Office.Interop.Outlook.AppointmentItem)objItem).UnRead)
                        blnUnreadMail = true;
                    else
                        if (!mlstEntryIDsViewed.Contains(((Microsoft.Office.Interop.Outlook.AppointmentItem)objItem).EntryID)) // email was not there when icon was last clicked to open outlook, could be an email that was not read by user, but flagged as read because the inbox folder is open and email is highlighted
                            blnUnreadMail = true;
                }
                else if (objItem is Microsoft.Office.Interop.Outlook.ContactItem)
                {
                    if (((Microsoft.Office.Interop.Outlook.ContactItem)objItem).UnRead)
                        blnUnreadMail = true;
                    else
                        if (!mlstEntryIDsViewed.Contains(((Microsoft.Office.Interop.Outlook.ContactItem)objItem).EntryID)) // email was not there when icon was last clicked to open outlook, could be an email that was not read by user, but flagged as read because the inbox folder is open and email is highlighted
                            blnUnreadMail = true;
                }
                else if (objItem is Microsoft.Office.Interop.Outlook.DistListItem)
                {
                    if (((Microsoft.Office.Interop.Outlook.DistListItem)objItem).UnRead)
                        blnUnreadMail = true;
                    else
                        if (!mlstEntryIDsViewed.Contains(((Microsoft.Office.Interop.Outlook.DistListItem)objItem).EntryID)) // email was not there when icon was last clicked to open outlook, could be an email that was not read by user, but flagged as read because the inbox folder is open and email is highlighted
                            blnUnreadMail = true;
                }
                else if (objItem is Microsoft.Office.Interop.Outlook.DocumentItem)
                {
                    if (((Microsoft.Office.Interop.Outlook.DocumentItem)objItem).UnRead)
                        blnUnreadMail = true;
                    else
                        if (!mlstEntryIDsViewed.Contains(((Microsoft.Office.Interop.Outlook.DocumentItem)objItem).EntryID)) // email was not there when icon was last clicked to open outlook, could be an email that was not read by user, but flagged as read because the inbox folder is open and email is highlighted
                            blnUnreadMail = true;
                }
                else if (objItem is Microsoft.Office.Interop.Outlook.JournalItem)
                {
                    if (((Microsoft.Office.Interop.Outlook.JournalItem)objItem).UnRead)
                        blnUnreadMail = true;
                    else
                        if (!mlstEntryIDsViewed.Contains(((Microsoft.Office.Interop.Outlook.JournalItem)objItem).EntryID)) // email was not there when icon was last clicked to open outlook, could be an email that was not read by user, but flagged as read because the inbox folder is open and email is highlighted
                            blnUnreadMail = true;
                }
                else if (objItem is Microsoft.Office.Interop.Outlook.MeetingItem)
                {
                    if (((Microsoft.Office.Interop.Outlook.MeetingItem)objItem).UnRead)
                        blnUnreadMail = true;
                    else
                        if (!mlstEntryIDsViewed.Contains(((Microsoft.Office.Interop.Outlook.MeetingItem)objItem).EntryID)) // email was not there when icon was last clicked to open outlook, could be an email that was not read by user, but flagged as read because the inbox folder is open and email is highlighted
                            blnUnreadMail = true;
                }
                else if (objItem is Microsoft.Office.Interop.Outlook.PostItem)
                {
                    if (((Microsoft.Office.Interop.Outlook.PostItem)objItem).UnRead)
                        blnUnreadMail = true;
                    else
                        if (!mlstEntryIDsViewed.Contains(((Microsoft.Office.Interop.Outlook.PostItem)objItem).EntryID)) // email was not there when icon was last clicked to open outlook, could be an email that was not read by user, but flagged as read because the inbox folder is open and email is highlighted
                            blnUnreadMail = true;
                }
                else if (objItem is Microsoft.Office.Interop.Outlook.RemoteItem)
                {
                    if (((Microsoft.Office.Interop.Outlook.RemoteItem)objItem).UnRead)
                        blnUnreadMail = true;
                    else
                        if (!mlstEntryIDsViewed.Contains(((Microsoft.Office.Interop.Outlook.RemoteItem)objItem).EntryID)) // email was not there when icon was last clicked to open outlook, could be an email that was not read by user, but flagged as read because the inbox folder is open and email is highlighted
                            blnUnreadMail = true;
                }
                else if (objItem is Microsoft.Office.Interop.Outlook.ReportItem)
                {
                    if (((Microsoft.Office.Interop.Outlook.ReportItem)objItem).UnRead)
                        blnUnreadMail = true;
                    else
                        if (!mlstEntryIDsViewed.Contains(((Microsoft.Office.Interop.Outlook.ReportItem)objItem).EntryID)) // email was not there when icon was last clicked to open outlook, could be an email that was not read by user, but flagged as read because the inbox folder is open and email is highlighted
                            blnUnreadMail = true;
                }
                else if (objItem is Microsoft.Office.Interop.Outlook.TaskItem)
                {
                    if (((Microsoft.Office.Interop.Outlook.TaskItem)objItem).UnRead)
                        blnUnreadMail = true;
                    else
                        if (!mlstEntryIDsViewed.Contains(((Microsoft.Office.Interop.Outlook.TaskItem)objItem).EntryID)) // email was not there when icon was last clicked to open outlook, could be an email that was not read by user, but flagged as read because the inbox folder is open and email is highlighted
                            blnUnreadMail = true;
                }
                else if (objItem is Microsoft.Office.Interop.Outlook.TaskRequestAcceptItem)
                {
                    if (((Microsoft.Office.Interop.Outlook.TaskRequestAcceptItem)objItem).UnRead)
                        blnUnreadMail = true;
                    else
                        if (!mlstEntryIDsViewed.Contains(((Microsoft.Office.Interop.Outlook.TaskRequestAcceptItem)objItem).EntryID)) // email was not there when icon was last clicked to open outlook, could be an email that was not read by user, but flagged as read because the inbox folder is open and email is highlighted
                            blnUnreadMail = true;
                }
                else if (objItem is Microsoft.Office.Interop.Outlook.TaskRequestDeclineItem)
                {
                    if (((Microsoft.Office.Interop.Outlook.TaskRequestDeclineItem)objItem).UnRead)
                        blnUnreadMail = true;
                    else
                        if (!mlstEntryIDsViewed.Contains(((Microsoft.Office.Interop.Outlook.TaskRequestDeclineItem)objItem).EntryID)) // email was not there when icon was last clicked to open outlook, could be an email that was not read by user, but flagged as read because the inbox folder is open and email is highlighted
                            blnUnreadMail = true;
                }
                else if (objItem is Microsoft.Office.Interop.Outlook.TaskRequestItem)
                {
                    if (((Microsoft.Office.Interop.Outlook.TaskRequestItem)objItem).UnRead)
                        blnUnreadMail = true;
                    else
                        if (!mlstEntryIDsViewed.Contains(((Microsoft.Office.Interop.Outlook.TaskRequestItem)objItem).EntryID)) // email was not there when icon was last clicked to open outlook, could be an email that was not read by user, but flagged as read because the inbox folder is open and email is highlighted
                            blnUnreadMail = true;
                }
                else if (objItem is Microsoft.Office.Interop.Outlook.TaskRequestUpdateItem)
                {
                    if (((Microsoft.Office.Interop.Outlook.TaskRequestUpdateItem)objItem).UnRead)
                        blnUnreadMail = true;
                    else
                        if (!mlstEntryIDsViewed.Contains(((Microsoft.Office.Interop.Outlook.TaskRequestUpdateItem)objItem).EntryID)) // email was not there when icon was last clicked to open outlook, could be an email that was not read by user, but flagged as read because the inbox folder is open and email is highlighted
                            blnUnreadMail = true;
                }
                else if (objItem is Microsoft.Office.Interop.Outlook.NoteItem)
                {
                    if (!mlstEntryIDsViewed.Contains(((Microsoft.Office.Interop.Outlook.NoteItem)objItem).EntryID)) // email was not there when icon was last clicked to open outlook, could be an email that was not read by user, but flagged as read because the inbox folder is open and email is highlighted
                        blnUnreadMail = true;
                }
                else
                    blnUnreadMail = true;   // unknown item, assume it's unread since there is no way to tell
            }
            mblnNotifyIconOn = blnUnreadMail;
        }

        void StartUpdateNotifyIconThread()
        {
            while (!mblnFormClosed)
            {
                if (mblnLoggedIn)
                {
                    if (mblnNotifyIconOn)
                    {
                        // show notification in system tray
                        objNotifyIcon.Icon = Properties.Resources.email_white;
                        Application.DoEvents();
                        Thread.Sleep(100);
                        if (!mblnFormClosed & mblnNotifyIconOn)
                        {
                            objNotifyIcon.Icon = Properties.Resources.email_blue;
                            Application.DoEvents();
                            Thread.Sleep(100);
                        }
                        if (!mblnFormClosed & mblnNotifyIconOn)
                        {
                            objNotifyIcon.Icon = Properties.Resources.email_green;
                            Application.DoEvents();
                            Thread.Sleep(100);
                        }
                        if (!mblnFormClosed & mblnNotifyIconOn)
                        {
                            objNotifyIcon.Icon = Properties.Resources.email_red;
                            Application.DoEvents();
                            Thread.Sleep(100);
                        }
                    }
                    else
                    {
                        // show regular icon in system tray
                        objNotifyIcon.Icon = Properties.Resources.eye;
                        Application.DoEvents();
                        Thread.Sleep(100);
                    }
                }
                else
                {
                    // show red icon to signify there was an error 
                    objNotifyIcon.Icon = Properties.Resources.email_red;
                    Application.DoEvents();
                    Thread.Sleep(100);
                }                
            }
        }

        private void objNotifyIcon_MouseClick(object sender, MouseEventArgs e)
        {
            switch(e.Button)
            {
                case MouseButtons.Right:
                    Show();
                    this.WindowState = FormWindowState.Normal;
                    break;

                case MouseButtons.Left:
                    // toggle between minimize and restore for microsoft outlook
                    Process[] pList = Process.GetProcessesByName("outlook");
                    if (pList.Length > 0) 
                    {
                        if (IsIconic(pList[0].MainWindowHandle))
                        {
                            // Remember All messages in inbox at moment of showing it
                            mlstEntryIDsViewed.Clear();
                            mlstEntryIDsViewed = new List<string>();
                            foreach (Object objItem in mobjOutlookInbox.Items)
                            {                                
                                if (objItem is Microsoft.Office.Interop.Outlook.MailItem)
                                    mlstEntryIDsViewed.Add(((Microsoft.Office.Interop.Outlook.MailItem)objItem).EntryID);
                                else if (objItem is Microsoft.Office.Interop.Outlook.AppointmentItem)
                                    mlstEntryIDsViewed.Add(((Microsoft.Office.Interop.Outlook.AppointmentItem)objItem).EntryID);
                                else if (objItem is Microsoft.Office.Interop.Outlook.ContactItem)
                                    mlstEntryIDsViewed.Add(((Microsoft.Office.Interop.Outlook.ContactItem)objItem).EntryID);
                                else if (objItem is Microsoft.Office.Interop.Outlook.DistListItem)
                                    mlstEntryIDsViewed.Add(((Microsoft.Office.Interop.Outlook.DistListItem)objItem).EntryID);
                                else if (objItem is Microsoft.Office.Interop.Outlook.DocumentItem)
                                    mlstEntryIDsViewed.Add(((Microsoft.Office.Interop.Outlook.DocumentItem)objItem).EntryID);
                                else if (objItem is Microsoft.Office.Interop.Outlook.JournalItem)
                                    mlstEntryIDsViewed.Add(((Microsoft.Office.Interop.Outlook.JournalItem)objItem).EntryID);
                                else if (objItem is Microsoft.Office.Interop.Outlook.MeetingItem)
                                    mlstEntryIDsViewed.Add(((Microsoft.Office.Interop.Outlook.MeetingItem)objItem).EntryID);
                                else if (objItem is Microsoft.Office.Interop.Outlook.PostItem)
                                    mlstEntryIDsViewed.Add(((Microsoft.Office.Interop.Outlook.PostItem)objItem).EntryID);
                                else if (objItem is Microsoft.Office.Interop.Outlook.RemoteItem)
                                    mlstEntryIDsViewed.Add(((Microsoft.Office.Interop.Outlook.RemoteItem)objItem).EntryID);
                                else if (objItem is Microsoft.Office.Interop.Outlook.ReportItem)
                                    mlstEntryIDsViewed.Add(((Microsoft.Office.Interop.Outlook.ReportItem)objItem).EntryID);
                                else if (objItem is Microsoft.Office.Interop.Outlook.TaskItem)
                                    mlstEntryIDsViewed.Add(((Microsoft.Office.Interop.Outlook.TaskItem)objItem).EntryID);
                                else if (objItem is Microsoft.Office.Interop.Outlook.TaskRequestAcceptItem)
                                    mlstEntryIDsViewed.Add(((Microsoft.Office.Interop.Outlook.TaskRequestAcceptItem)objItem).EntryID);
                                else if (objItem is Microsoft.Office.Interop.Outlook.TaskRequestDeclineItem)
                                    mlstEntryIDsViewed.Add(((Microsoft.Office.Interop.Outlook.TaskRequestDeclineItem)objItem).EntryID);
                                else if (objItem is Microsoft.Office.Interop.Outlook.TaskRequestItem)
                                    mlstEntryIDsViewed.Add(((Microsoft.Office.Interop.Outlook.TaskRequestItem)objItem).EntryID);
                                else if (objItem is Microsoft.Office.Interop.Outlook.TaskRequestUpdateItem)
                                    mlstEntryIDsViewed.Add(((Microsoft.Office.Interop.Outlook.TaskRequestUpdateItem)objItem).EntryID);
                                else if (objItem is Microsoft.Office.Interop.Outlook.NoteItem)
                                    mlstEntryIDsViewed.Add(((Microsoft.Office.Interop.Outlook.NoteItem)objItem).EntryID);
                            }

                            // show outlook                            
                            ShowWindow(pList[0].MainWindowHandle, SW_SHOWNOACTIVATE);
                            SetForegroundWindow(pList[0].MainWindowHandle);                            
                        }
                        else
                        {
                            // minimize outlook
                            ShowWindow(pList[0].MainWindowHandle, SW_MINIMIZE);
                        }
                    } 

                    // set the icon flag, don't wait a few seconds till it happens automatically
                    SetIconFlag();

                    break;
            }
        }

        private void frmMain_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                Hide();
            }
        }

        private void frmMain_FormClosed(object sender, FormClosedEventArgs e)
        {
            mblnFormClosed = true;
        }
    }
}
