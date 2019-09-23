using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit;
using MimeKit;
using System.IO;
using System.Globalization;
using System.Printing;
using System.Diagnostics;
//using Ghostscript.NET.Processor;
using PdfiumViewer;
using System.Drawing.Printing;
using System.Drawing;
using System.Threading.Tasks;
using System.Threading;

namespace WpfPrintFromImap
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 
    class MailSnippet
    {

        //date to pack order
        readonly private DateTime packing_day;
        //number of pages to print(send to printer)
        readonly private uint no_of_pages;
        //this string contains 3 terms: "'order' order_no city".
        readonly private string order_number;
        readonly private string searchValue;
        readonly private string attachmentName;
        readonly private string mailBody;
        readonly private string subject;
        
        public string getSearchValue()
        {
            return this.searchValue;
        }
        public DateTime getPackingDay()
        {
            return this.packing_day;
        }
        public uint getNoOfPages()
        {
            return this.no_of_pages;
        }
        public string getOrderNumber()
        {
            return this.order_number;
        }
        public string getAttachmentName()
        {
            return this.attachmentName;
        }
        public string getMailBody()
        {
            return this.mailBody;
        }
        public string getSubject()
        {
            return this.subject;
        }
        /// <summary>
        /// This function will try to filter out some of the most common variations of the subject, that has impact on the routine.
        /// Problem arise if mail is forwarded or replied to and some senders adding messages in the subject-line. There's alot of variations,
        /// so I'll just have to add them as they come.
        /// </summary>
        /// <returns>Mail-subject in a List of strings</returns>
        public List<String> FilterSubject(String str)
        {
            List<String> T = new List<String>();
            string[] table = str.Split(' ');
            
            string prevWord = null;
            foreach (string word in table)
            {
                if ((word.ToString() == "-") ||
                    (word.ToString().ToLower() == "fw:") || //forwarded message
                    (word.ToString().ToLower() == "re:") || //replied to message
                    (word.ToString() == ""))
                    continue;
                else
                {
                    if (prevWord != null)
                        if (prevWord.ToString().ToLower() == "ark") //someone is trying to add some kind of msg after the standard subject..we don't want it..break off.
                            break;
                    T.Add(word);
                    //keep track of previous word so we can look for certain placeholders
                    prevWord = word;
                }
            }
            return T;
        }
        /// <summary>
        /// Fills in private members of this class. Fills in info for one mail. For example date, number of attachments to print...etc
        /// </summary>
        /// <param name="subj">Subject of the mail</param>
        /// <param name="bde">Bodypart of the mail</param>
        /// <param name="attachName">name of attachment</param>
        public MailSnippet(string subj, string bde, string attachName)
        {
            subject = subj;
            List<String> str = FilterSubject(subj);

            CultureInfo cultureInfo = new CultureInfo("nb-NO");
            no_of_pages = 0;
            try
            {
                no_of_pages = uint.Parse(str[5]);
            }
            catch (FormatException)
            {
                //invalid numeric expression. we should concatenate index 4 & 5 into index 4 and delete index 5
                str[4] = string.Concat(str[4], " ", str[5]);
                str.Remove(str[5]);
            }

            try
            {
                packing_day = DateTime.ParseExact(str[1], "ddMMyy", cultureInfo);
            }
            catch (FormatException)
            {
                try
                {
                    packing_day = DateTime.ParseExact(str[1], "ddMMyyyy", cultureInfo); //just try..maybe it will work.
                }
                catch (FormatException ex)
                {
                    MessageBox.Show($"Cannot find date format of this header: {ex}. Please remove mail with order number {str[3]}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            order_number = str[2] + " " + str[3];
            searchValue = string.Concat("Pakkedag ", str[1]);
            mailBody = bde;
            attachmentName = attachName;
        }
    }
    public partial class MainWindow : Window
    {
        //private string
        private const string fileSettings = "settings.cfg";
        //data in file is in the same order as in the settigs dialogbox
        //1. imap-server
        readonly private string imap_server;
        //2. username
        readonly private string imap_user;
        //3. password //should be encrypted somehow(not in clear text at least)
        readonly private string imap_user_password;
        readonly string printer_adhesive;
        readonly string printer_plain;
        public string packing_day;
        List<MailSnippet> mailSnippets;
        readonly private string att_dir;
        public SearchQuery query { get; private set; }

        public string GetFilename()
        {
            return MainWindow.fileSettings;
        }
        public string getAttachmentDirectory()
        {
            return this.att_dir;
        }
        public MainWindow()
        {
            InitializeComponent();
            if(!File.Exists(GetFilename()))
            {
                MessageBox.Show("You must create a Settingsfile before continuing. Go to \"Settings\" and fill in the form.", "Settings are missing", MessageBoxButton.OK, MessageBoxImage.Hand);
            }
            else
            {
                try
                {
                    StreamReader file = new StreamReader(GetFilename());
                    this.imap_server = file.ReadLine();
                    this.imap_user = file.ReadLine();
                    string line = file.ReadLine();
                    if (line != null) //try to do this ..bla bla bla
                    {
                        var pswd = System.Convert.FromBase64String(line);
                        this.imap_user_password = System.Text.Encoding.UTF8.GetString(pswd);
                    }
                    this.printer_adhesive = file.ReadLine();
                    this.printer_plain = file.ReadLine();
                    this.packing_day = file.ReadLine();
                    this.txtPackingDay.Text += this.packing_day;
                    file.Close();
                }
                catch (Exception e)
                {
                    MessageBox.Show("Unhandled exceptions in reading settings-file: {0}", e.ToString(), MessageBoxButton.OK, MessageBoxImage.Error );
                    //the file exists, but is corrupted somehow. Delete it.
                    if(File.Exists(GetFilename()))
                    {//delete the file and let's try again.
                        
                        File.Delete(GetFilename());
                    }
                }

            }
            att_dir = "Attachments";
            Directory.CreateDirectory(att_dir);
            mailSnippets = new List<MailSnippet>();
        }

       

        public void mailSnippets_RemoveAll()
        {
            if(mailSnippets.Count > 0)
            {
                mailSnippets.Clear();
            }
        }
        //Exit the application
        private void MenuItem_Exit(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }
        //Open Messagebox that says a little about the application
        private void MenuItem_About(object sender, RoutedEventArgs e)
        {
            string msgAbout ="Written and designed by Frode Meek.\n" +
                "Only intended for use by Beco Lager AS and Nordkak AS";
            string msgCaption = "About this Application";
            MessageBoxButton button = MessageBoxButton.OK;
            MessageBoxImage icon = MessageBoxImage.Information;
            MessageBox.Show(msgAbout, msgCaption, button, icon);

        }
        //Open dialogbox that show curent setings that are Being used
        private void MenuItem_Settings(object sender, RoutedEventArgs e)
        {
            try
            {
                Window1 winTest = new Window1();
                winTest.Owner = this;
                winTest.Show();
                
            }
            catch(Exception ex)
            {
                MessageBox.Show("Unhandled exception has occured: " + ex.Message, "Exception", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        //When this button is pressed there are a couple of  questions that has to be answered:
        // 1. Are we updating? In That case: add new list-item
        // 2. Are we making a new list? Is it a new date?
        // -------------------------------------------------
        //
        private void BtnUpdateMailList_Click(object sender, RoutedEventArgs e)
        {
            var client = new ImapClient();
            bool bFail = false;
            // For demo-purposes, accept all SSL certificates
            try
            {
                client.ServerCertificateValidationCallback = (s, c, h, ex) => true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Caught an exception under ServerCertificateValidationCallback: " + ex.Message, "Unhandled Exception Cought", MessageBoxButton.OK, MessageBoxImage.Error);
                bFail = true;
            }

            try
            {
                client.Connect(this.imap_server, 993, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Caught an exception under Connect: " + ex.Message, "Unhandled Exception Cought", MessageBoxButton.OK, MessageBoxImage.Error);
                bFail = true;
            }
            try
            {
                client.Authenticate(this.imap_user, this.imap_user_password);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Caught an exception under Authenticate: " + ex.Message, "Unhandled Exception Cought", MessageBoxButton.OK, MessageBoxImage.Error);
                bFail = true;
            }
            if (bFail)//no reason to continue anymore
                return;
            // The Inbox folder is always available on all IMAP servers...
            var inbox = client.Inbox;

            inbox.Open(FolderAccess.ReadOnly);
            string text = this.packing_day;
           
            var query = SearchQuery.SubjectContains(text);
            //IList<UniqueId> T = new List<UniqueId>();
            //temporary list that has to be compared with MainWindow-membervariable mailSnippets
            //Connection to imap-server is good. Removeall in list and populate with new list
            this.lstBxMails.Items.Clear();
            List<MailSnippet> tempSnippets = new List<MailSnippet>();
            var orderBy = new[] { OrderBy.Arrival };
            foreach (var uid in inbox.Sort(query, orderBy))
            {
                var message = inbox.GetMessage(uid);
                foreach (var attachment in message.Attachments)
                {
                    var fileName = attachment.ContentDisposition.FileName;
                    MailSnippet ms = new MailSnippet(message.Subject, message.TextBody, fileName);
                    tempSnippets.Add(ms);
                    using (var stream = File.Create(this.att_dir + "\\" + fileName))
                    {
                        if (attachment is MessagePart)
                        {
                            var rfc822 = (MessagePart)attachment;

                            rfc822.Message.WriteTo(stream);
                        }
                        else
                        {
                            var part = (MimePart)attachment;
                            part.Content.DecodeTo(stream);
                        }
                    }
                }
            }
            //if number of items is 0 then we can just add the list(if any)
            //this will always be true since I remove everything in the listbox when the button is pushed
            //easy/simple solution instead of comparing objects and lists.
            if (this.lstBxMails.Items.Count == 0)
            {
                mailSnippets = tempSnippets;
                foreach(MailSnippet mailSubj in mailSnippets)
                {
                    lstBxMails.Items.Add(mailSubj.getSubject());
                }
            }
            tempSnippets = null;
            client.Disconnect(true); 
        }

        private void LstBxMails_SelectionChanged(object sender, RoutedEventArgs e)
        {
            var index = this.lstBxMails.SelectedIndex;
            if (index >= 0)
            {
                this.txtBxAttachment.Text = mailSnippets[index].getAttachmentName();
                this.txtMailBody.Text = mailSnippets[index].getMailBody();
            }
        }
        //bad solution...Using GhostScript..We will try using "direct print" by sending th epdf directly to the queue
/*
        public void PrintPDF(string printer, string paperName, string filename, int copies)
        {
            using (GhostscriptProcessor processor = new GhostscriptProcessor())
            {
                List<string> switches = new List<string>();
                switches.Add("-empty");
                switches.Add("-dPrinted");
                switches.Add("-dBATCH");
                switches.Add("-dNOPAUSE");
                switches.Add("-dNOSAFER");
                switches.Add("-dNumCopies="+ copies.ToString());
                switches.Add("-sDEVICE=mswinpr2");
                switches.Add("-sOutputFile=%printer%" + printer);
                switches.Add("-f");
                switches.Add(filename);

                processor.StartProcessing(switches.ToArray(), null);
            }

        }a
  */
        private void BtnPrint_Click(object sender, RoutedEventArgs e)
        {
            string current_path = Directory.GetCurrentDirectory();

            string fullpath = current_path + "\\" + att_dir + "\\";
            // string file = "test.pdf";

            foreach (MailSnippet ms in mailSnippets)
            {
                string filename = fullpath + ms.getAttachmentName();
                string txtFile = ms.getSubject() + ".txt";
                
                PrintDocument pdt = new System.Drawing.Printing.PrintDocument();
                pdt.DocumentName = ms.getSubject();
                pdt.PrinterSettings.PrinterName = printer_plain;
                pdt.PrinterSettings.Copies = 1;
                pdt.PrintPage += delegate (object sender1, PrintPageEventArgs e1)
                {
                    e1.Graphics.DrawString(ms.getMailBody(), new Font("Times New Roman", 20), new SolidBrush(Color.Black), new RectangleF(0, 0, pdt.DefaultPageSettings.PrintableArea.Width, pdt.DefaultPageSettings.PrintableArea.Height));
                };
                pdt.Print();
                Action<object> printAttachment = (object obj) =>
                {
                    using (var document = PdfDocument.Load(filename))
                    {
                        using (var printDocument = document.CreatePrintDocument())
                        {
                            printDocument.PrinterSettings.PrintFileName = filename;
                            printDocument.PrinterSettings.PrinterName = printer_adhesive;
                            printDocument.DocumentName = filename;
                            printDocument.PrinterSettings.PrintFileName = filename;
                            printDocument.PrinterSettings.Copies = (short)ms.getNoOfPages();
                            printDocument.PrintController = new System.Drawing.Printing.StandardPrintController();
                            printDocument.Print();

                        }
                    }
                };
                Task pa = new Task(printAttachment, "stopAttachment");
                pa.Start();
                pa.Wait();
                Thread.Sleep(2000);
            }
                 
        }
        
        private void BtnRemoveMail_Click(object sender, RoutedEventArgs e)
        {
            int index = this.lstBxMails.SelectedIndex;
            if(index < 0)
            {
                MessageBox.Show("You must select an item to delete it.", "Warning", MessageBoxButton.OK, MessageBoxImage.Hand);
            }
            else
            {
                this.lstBxMails.Items.Remove(lstBxMails.SelectedItems[0]);
                this.mailSnippets.RemoveAt(index);
                this.lstBxMails.UpdateLayout();
            }
        }

        private void BtnRemoveAllMails_Click(object sender, RoutedEventArgs e)
        {
            int items = lstBxMails.Items.Count;
            if (items > 0)
            {
                lstBxMails.Items.Clear();
                mailSnippets.Clear();
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            //when we close the application we can delete all files in the "Attachments" directory
            string current_path = Directory.GetCurrentDirectory();

            string fullpath = current_path + "\\" + att_dir + "\\";
            try
            {
                Directory.Delete(fullpath, true);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
    }
}
