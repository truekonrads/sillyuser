using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Threading;

namespace SillyUser
{

    public partial class Form1 : Form
    {
        private Outlook.Application app;
        private Outlook.NameSpace ns;        
        public Form1()
        {
            app = new Outlook.Application();
            ns = app.Session;
            InitializeComponent();
            var textboxglue = new TextBoxStreamWriter(this.richTextBox1,this);            
            var t = new TextWriterTraceListener(textboxglue) { TraceOutputOptions = TraceOptions.DateTime };
            Trace.Listeners.Add(t);
            Trace.AutoFlush = true;
            Trace.TraceInformation("Form1 initialised");
            new Thread(() =>
            {
                Thread.CurrentThread.IsBackground = true;
                StartBeingSilly();
            }).Start();
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
            var inbox = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Trace.TraceInformation($"Dum di dum, reading {inbox.FullFolderPath}!");
            foreach (Outlook.MailItem item in inbox.Items)
            {
                if (item.UnRead)
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
                    var sleeptime = 5000;
                    Trace.TraceInformation($"Sleeping for {sleeptime}");
                    Thread.Sleep(sleeptime);
                }
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

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
