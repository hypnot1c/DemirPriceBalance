using System;
using System.Collections.Generic;
using System.ComponentModel;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

using DemirPriceBalance.Logic;

namespace DemirPriceBalance
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

    private void btnMerge_Click(object sender, RoutedEventArgs e)
    {
      prbWork.IsIndeterminate = true;
      var wrk = new BackgroundWorker();
      wrk.DoWork += worker_DoWork;
      wrk.ProgressChanged += worker_ProgressChanged;
      wrk.RunWorkerCompleted += worker_RunWorkerCompleted;
      wrk.RunWorkerAsync();
    }

    private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
      //throw new NotImplementedException();
      prbWork.IsIndeterminate = false;
    }

    private void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
    {
      throw new NotImplementedException();
    }

    private void worker_DoWork(object sender, DoWorkEventArgs e)
    {
      ExcelReader.readExcel();
      //throw new NotImplementedException();
    }
  }
}
