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
    /// Interaction logic for WinProgress.xaml
    /// </summary>
    /// 
    
    public partial class WinProgress : Window
    {
        //this is probably very unsafe, but it's a solution for the moment..MWobj is a public object of the main window. this is sent back so we can update 
        //this dialogs' text-box from the main window on another thread isung Dispatcher.Invoke()....
        private async void StartProgress(MainWindow MWobj)
        {
            ProgressBar.IsIndeterminate = true;

            await Task.Run(() =>
            {
                MWobj.OpenConnectMails(this);
            
            });
            MWobj.PopulateListBox();
            this.Close();
        }
        
        public WinProgress(MainWindow MWobj)
        {
            InitializeComponent();
            StartProgress(MWobj);
        }
    }
}
