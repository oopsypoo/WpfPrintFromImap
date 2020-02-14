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
using CredentialManagement;


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
                    client.ServerCertificateValidationCallback = (s, c, h, ex) => true;
                }
                catch (Exception e)
                {
                    MessageBox.Show("Caught an exception under ServerCertificateValidationCallback: " + e.Message, "Unhandled Exception Cought", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }

                try
                {
                    string imap_server = this.txtBxImapServer.Text;
                    client.SslProtocols = System.Security.Authentication.SslProtocols.Default;
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
                    
                    client.Authenticate(username, CredentialUtil.GetCredentials(((MainWindow)Application.Current.MainWindow).GetCredentialName()));
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

            this.GetAvailablePrinters();
            bool bFileCorrupted = false;
            if (File.Exists(AppConfig.SettingsFile))
            {
                AppConfig appConfig = ((MainWindow)Application.Current.MainWindow).ac;
                this.txtBxImapServer.Text = appConfig.MailServer;
                if (this.txtBxImapServer.Text == null)
                    bFileCorrupted = true;
                this.txtBxUserName.Text = appConfig.Login;
                if (this.txtBxUserName.Text == null)
                    bFileCorrupted = true;


                /* do something abvout this...not good
                 * 
                 */
                //Get from vault
                //this.txtBxPassword.Password = System.Text.Encoding.UTF8.GetString(pswd);
                //fix this...not returning the credentials
                this.txtBxPassword.Password = CredentialUtil.GetCredentials(((MainWindow)Application.Current.MainWindow).GetCredentialName());
                
                this.lstBxPrinterAdhesiveLabel.SelectedValue = appConfig.AdhessivePrinter;
/*                if (this.lstBxPrinterAdhesiveLabel.SelectedValue == null)
                    bFileCorrupted = true;
 */               this.lstBxPrinterPlain.SelectedValue = appConfig.StandardPrinter;
/*                if (this.lstBxPrinterPlain.SelectedValue == null)
                    bFileCorrupted = true;
  */              this.packingday = appConfig.PackingDay;
                if (this.packingday == null)
                    bFileCorrupted = true;
                this.txtBxMailFilterSubject.Text = packingday;
                if (bFileCorrupted)
                {
                    MessageBox.Show("Please try to set up Credentials again..", "File-corruption", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
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
            if (((MainWindow)Application.Current.MainWindow).ac == null)
                ((MainWindow)Application.Current.MainWindow).ac = new AppConfig();
            AppConfig appConfig = ((MainWindow)Application.Current.MainWindow).ac;
            
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
            appConfig.Login = txtBxUserName.Text;
            appConfig.MailServer = txtBxImapServer.Text;
            appConfig.StandardPrinter = lstBxPrinterPlain.SelectedItem.ToString();
            appConfig.AdhessivePrinter = lstBxPrinterAdhesiveLabel.SelectedItem.ToString();
            appConfig.PackingDay = this.packingday;

            appConfig.SerializeAppConfig();
            if (!bError)
            {
                CredentialUtil.SetCredentials(((MainWindow)Application.Current.MainWindow).GetCredentialName(), null, txtBxPassword.Password, PersistanceType.LocalComputer);
                
                if (this.date_changed)
                {
                    var mw = ((MainWindow)Application.Current.MainWindow);
                    mw.lstBxMails.Items.Clear();
                    mw.mailSnippets_RemoveAll();
                    mw.txtPackingDay.Text = "Gjeldende: " + packingday;
                }
                //close window
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
            if (lstBxPrinterAdhesiveLabel.SelectedItem.ToString() != null)
                ((MainWindow)Application.Current.MainWindow).SetPrinterAdhessive(lstBxPrinterAdhesiveLabel.SelectedItem.ToString());
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
            if(lstBxPrinterPlain.SelectedItem.ToString() != null)
                ((MainWindow)Application.Current.MainWindow).SetPrinterPlain(lstBxPrinterPlain.SelectedItem.ToString());
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
                ((MainWindow)Application.Current.MainWindow).ac.PackingDay = this.packingday;
        }
    }
}
