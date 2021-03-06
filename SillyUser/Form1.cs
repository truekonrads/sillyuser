﻿using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;
using Outlook = NetOffice.OutlookApi;
using System.Diagnostics;
using System.Threading;
using NetOffice.OutlookApi.Enums;
namespace SillyUser
{

    public partial class Form1 : Form
    {
        private Outlook.Application app;
        private Outlook._NameSpace ns;        
        private int  sleeptime = 5000;
        Outlook.MAPIFolder currentFolder = null;

        private List<Outlook.MAPIFolder> GetInboxes()
        {
            var l = new List<Outlook.MAPIFolder>();
            foreach (Outlook.Store store in ns.Stores)
            {
                l.Add(store.GetDefaultFolder(OlDefaultFolders.olFolderInbox));
            }
            return l;
        }

        private class Item
        {
            
            public Outlook.MAPIFolder Value;
            public Item(Outlook.MAPIFolder value)
            {
                Value = value;
            }
            public override string ToString()
            {
                // Generates the text shown in the combo box
                return Value.FolderPath;
            }
        }

        public Form1()
        {

            InitializeComponent();
            var textboxglue = new TextBoxStreamWriter(this.richTextBox1,this);            
            var t = new TextWriterTraceListener(textboxglue) { TraceOutputOptions = TraceOptions.DateTime };
            


            Trace.Listeners.Add(t);
            Trace.AutoFlush = true;
            app = new Outlook.Application();
            ns = app.Session;
            var inboxes = GetInboxes();
            
            foreach (var i in inboxes)
                { comboBox1.Items.Add(new Item(i)); }
            comboBox1.SelectedItem = comboBox1.Items[0];
            Trace.TraceInformation("Form1 initialised");

        }
        private List<string> FindUrlsInText(string text){
            List<string> links = new List<string>();
            char[] delimiters = new char[] { ' ', '\t', '\n', '\r' };
            foreach (string chunk in text.Split(delimiters)) {
                if (chunk.IndexOf(@"http://", StringComparison.OrdinalIgnoreCase) == 0 ||
                    chunk.IndexOf(@"https://", StringComparison.OrdinalIgnoreCase) == 0)
                {
                    links.Add(chunk);
                }
            }
            return links;
        }
        private void ProcessOneItem(Outlook.MailItem item)
        {

            Trace.TraceInformation($"Processing email from {item.SenderEmailAddress} \"{ item.Subject}\" sent on  {item.SentOn}");
            foreach (Outlook.Attachment att in item.Attachments)
            {
                
                int epoch = (int)(DateTime.UtcNow - new DateTime(1970, 1, 1)).TotalSeconds;
                var tempAttPath = Environment.ExpandEnvironmentVariables($"%TEMP%\\{epoch}_{att.FileName}");
                att.SaveAsFile(tempAttPath);
                Start(tempAttPath);
                return;
            }
            var links = FindUrlsInText(item.Body);
            foreach(var l in links)
            {
                Start(l);
            }

        }

        private static void Start(string tempAttPath)
        {
            ProcessStartInfo psi = new ProcessStartInfo(tempAttPath)
            {
                UseShellExecute = true
            };
            Process.Start(psi);
        }

        private void EmailLoop()
        {

            Trace.TraceInformation("Started e-mail loop");

            var inbox = currentFolder;
            Trace.TraceInformation($"Dum di dum, reading {inbox.FullFolderPath}!");
            foreach (Outlook.MailItem item in inbox.Items)
            {
                if (item.UnRead && item.Body!=null)
                {
                    ProcessOneItem(item);
                    item.UnRead = false;
                }
            }
            Trace.TraceInformation("Done with email loop");
        }
        public void StartBeingSilly()
        {
            while (true)
            {
                try
                {
                    EmailLoop();
                }
                catch (Exception e)
                {
                    var trace = new StackTrace(e, true);
                    Trace.TraceError(e.ToString());
                    Trace.TraceError(trace.ToString());
                }
                finally
                {
                    //var sleeptime = 5000;
                    Trace.TraceInformation($"Sleeping for {sleeptime}");
                    Thread.Sleep(sleeptime);
                }
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            richTextBox1.SelectionStart = richTextBox1.Text.Length;
            // scroll it automatically
            richTextBox1.ScrollToCaret();
        }

        private void textBox1_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {

            try
            {
                var t=Convert.ToInt16(textBox1.Text);
                if (t > 0)
                {
                    sleeptime = t;
                    errorProvider1.Clear();
                    errorProvider1.SetError(textBox1, "");
                }
                else
                {
                    errorProvider1.SetError(textBox1, "Just unsigned integers please");
                }
            }
            catch 
            {
                errorProvider1.SetError(textBox1,"Just unsigned integers please");
            }
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            var item = ((ComboBox)sender).SelectedItem;
            currentFolder = ((Item) item).Value;
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            new Thread(() =>
            {
                Thread.CurrentThread.IsBackground = true;
                StartBeingSilly();
            }).Start();
        }
    }


    public class TextBoxStreamWriter : TextWriter
    {
        private Boolean _disposing;
        private RichTextBox _textBox;
        private Form form;
        
        public override Encoding Encoding {
            get {
                return  new System.Text.UTF8Encoding();
            }
        }

        public TextBoxStreamWriter(RichTextBox textBox, Form f)        
        {
            _textBox = textBox;
            form = f;
        }

        public override void WriteLine()
        {
            WriteLine(Environment.NewLine);
        }

        delegate void WriteLineCalback(string text);
        public override void WriteLine(string value)
        {

            if (_textBox.IsDisposed)
            {
                return;
            }
            if (_textBox.InvokeRequired)
            {
                WriteLineCalback d = new WriteLineCalback(WriteLine);
                form.Invoke(d, new object[] { value});
            }
            else
            {
                
                _textBox.AppendText(value + Environment.NewLine);
            }
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);

            _disposing = disposing;
        }



    }
    
}
