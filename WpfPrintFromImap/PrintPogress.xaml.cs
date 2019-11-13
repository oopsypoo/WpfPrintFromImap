using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using WpfPrintFromImap;


namespace WpfPrintFromImap
{
    /// <summary>
    /// Interaction logic for PrintPogress.xaml
    /// </summary>
    public partial class PrintProgress : Window
    {
        private async void StartProgress(MainWindow MWobj)
        {
            this.printProgress.Maximum = MWobj.lstBxMails.Items.Count;
            this.printProgress.Value = 0;
            this.ppText.Text = "Starting printing of mails.";
            await Task.Run(() =>
            {
                MWobj.SendToPrinter(this);
                   
            });
            this.Close();
        }
        public PrintProgress(MainWindow MWin)
        {
            
            InitializeComponent();
            StartProgress(MWin);
        }
    }
}
