using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit;
using MimeKit;
using System.IO;
using Microsoft.Win32;
using System.Drawing;

namespace WpfPrintFromImap
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// IMAP and Printer settings for this appication to work
    /// IMAP -server and Imap username
    /// other settings are automated. Connection should be SSL/TLS (for example)
    /// </summary>
    public partial class Window1 : Window
    {
        const string filename = "settings.cfg";
        //this is the packingday var that will be saved to file...easier this way
        string packingday;
        bool date_changed = false;
        private void GetAvailablePrinters()
        {
            int index = 0;
            foreach (string printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            {
                lstBxPrinterAdhesiveLabel.Items.Insert(index, printer);
                lstBxPrinterPlain.Items.Insert(index, printer);
                index++;
            }
        }
        private bool CheckIMAPConnection()
        {
            using (var client = new ImapClient())
            {
                // For demo-purposes, accept all SSL certificates
                try
                {
                    client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                }
                catch (Exception e)
                {
                    MessageBox.Show("Caught an exception under ServerCertificateValidationCallback: " + e.Message, "Unhandled Exception Cought", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }

                try
                {
                    string imap_server = this.txtBxImapServer.Text;
                    client.Connect(imap_server, 993, true);
                }
                catch (Exception ex)
                {
                    if (ex is ImapCommandException)
                    {
                        return false;
                    }
                    else if (ex is ImapProtocolException)
                    {
                        return false;
                    }
                    else
                    {
                        MessageBox.Show("Caught an exception under Connect: " + ex.Message, "Unhandled Exception Cought", MessageBoxButton.OK, MessageBoxImage.Error);
                        return false;
                    }
                }
                try
                {
                    string username = this.txtBxUserName.Text;
                    string password = this.txtBxPassword.Password;
                    client.Authenticate(username, password);
                    password = null;
                }
                catch (Exception e)
                {
                    MessageBox.Show("Caught an exception under Authenticate: " + e.Message, "Unhandled Exception Cought", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
                client.Disconnect(true);
                return true;
            }
        }

        public Window1()
        {
            InitializeComponent();
            bool bFileCorrupted = false;
            if (File.Exists(filename))
            {
                StreamReader file = new StreamReader(filename);
                this.txtBxImapServer.Text = file.ReadLine();
                if (this.txtBxImapServer.Text == null)
                    bFileCorrupted = true;
                this.txtBxUserName.Text = file.ReadLine();
                if (this.txtBxUserName.Text == null)
                    bFileCorrupted = true;
                string line = file.ReadLine();
                if (line != null)
                {
                    var pswd = System.Convert.FromBase64String(line);
                    this.txtBxPassword.Password = System.Text.Encoding.UTF8.GetString(pswd);
                }
                else
                    bFileCorrupted = true;
                this.GetAvailablePrinters();
                this.lstBxPrinterAdhesiveLabel.SelectedValue = file.ReadLine();
/*                if (this.lstBxPrinterAdhesiveLabel.SelectedValue == null)
                    bFileCorrupted = true;
 */               this.lstBxPrinterPlain.SelectedValue = file.ReadLine();
/*                if (this.lstBxPrinterPlain.SelectedValue == null)
                    bFileCorrupted = true;
  */              this.packingday = file.ReadLine();
                if (this.packingday == null)
                    bFileCorrupted = true;
                this.txtBxMailFilterSubject.Text = packingday;
                if (bFileCorrupted)
                {
                    MessageBox.Show("Settingsfile is corrupted...deleting. Please set up your settings again and save it.", "File-corruption", MessageBoxButton.OK, MessageBoxImage.Warning);
                    file.Close();
                    File.Delete(filename);
                }
                file.Close();
            }

            if(string.IsNullOrEmpty(packingday))
            {
                this.calPakkedag.SelectedDate = DateTime.Now;
                this.txtBxMailFilterSubject.Text = "Pakkedag " + calPakkedag.SelectedDate.Value.ToString("ddMMyy");
                this.packingday = txtBxMailFilterSubject.Text;
            }
        }
        //close current window and save data if it allready has not been done
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            bool bError = false;
            //Save data to file from text-boxes and save data to file
            if (txtBxImapServer.Text == null)
            {
                MessageBox.Show("IMAP servername cannot be empty", "Error", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                bError = true;
            }
            if(txtBxUserName.Text == null)
            { 
                MessageBox.Show("Username cannot be empty", "Error", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                bError = true;
            }
            if (txtBxPassword.Password == null)
            { 
                MessageBox.Show("Password cannot be empty", "Error", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                bError = true;
            }
            //if there are no errors save data and close dialog
            if (!bError)
            {
                StreamWriter file = new StreamWriter(filename);
                file.WriteLine(this.txtBxImapServer.Text);
                file.WriteLine(this.txtBxUserName.Text);
                //just convert to base64 for now..."not clear"
                var text = System.Text.Encoding.UTF8.GetBytes(this.txtBxPassword.Password);
                file.WriteLine(Convert.ToBase64String(text));
                if (lstBxPrinterAdhesiveLabel.SelectedItem == null)
                    file.WriteLine((char)0x00);
                else
                    file.WriteLine(lstBxPrinterAdhesiveLabel.SelectedItem.ToString());
                if (lstBxPrinterPlain.SelectedItem == null)
                    file.WriteLine((char)0x00);
                else
                    file.WriteLine(lstBxPrinterPlain.SelectedItem.ToString());
                file.WriteLine(this.packingday);
                file.Close();
                //if date has changed we can remove elements from the listbox in the main window
                //and delete whatever is in it
                if (this.date_changed)
                {
                    var mw = ((MainWindow)Application.Current.MainWindow);
                    mw.lstBxMails.Items.Clear();
                    mw.mailSnippets_RemoveAll();
                    mw.txtPackingDay.Text = "Gjeldende: " + packingday;
                }
                //clsoe window
                this.Close();
            }
        }
        private void Button_TestConnection(object sender, RoutedEventArgs e)
        {
            if(!this.CheckIMAPConnection())
            {
                MessageBox.Show("Could Not connect. Check you settings", "Connection-Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
                MessageBox.Show("Settings are ok", "Connection", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void BtnChangeAdhesiveLabelPrinter_Click(object sender, RoutedEventArgs e)
        {
 /*           MessageBox.Show(lstBxPrinterAdhesiveLabel.SelectedItem.ToString());
            if (MessageBox.Show("Do you want to append this to the file?", "Warning! Appending data to file", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                StreamWriter file = new StreamWriter(filename, append: true);
                file.WriteLine(lstBxPrinterAdhesiveLabel.SelectedItem.ToString());
                file.Close();
            }
 */

    }

        private void BtnChangePlainA4Printer_Click(object sender, RoutedEventArgs e)
        {
 /*           MessageBox.Show(lstBxPrinterPlain.SelectedItem.ToString());
            if (MessageBox.Show("Do you want to append this to the file?", "Warning! Appending data to file", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                StreamWriter file = new StreamWriter(filename, append: true);
                file.WriteLine(lstBxPrinterPlain.SelectedItem.ToString());
                file.Close();
            }
   */
            
        }

        private void CalPakkedag_SelectedDatesChanged(object sender, SelectionChangedEventArgs e)
        {
            var calendar = sender as Calendar;

            // ... See if a date is selected.
            if (calendar.SelectedDate.HasValue)
            {
                // ... Display SelectedDate in Title.
                DateTime date = calendar.SelectedDate.Value;
                this.txtBxMailFilterSubject.Text = "Pakkedag " + date.ToString("ddMMyy");
                this.packingday = txtBxMailFilterSubject.Text;
                this.date_changed = true;
            }
        }

        private void IMAP_Settings_Closed(object sender, EventArgs e)
        {
            if(this.packingday != null)
                ((MainWindow)Application.Current.MainWindow).packing_day = this.packingday;
        }
    }
}
