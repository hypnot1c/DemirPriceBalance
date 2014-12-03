using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
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
//using System.Windows.Shapes;

using DemirPriceBalance.Logic;
using Microsoft.Win32;
using System.Diagnostics;

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
      txtDemirTires.Text = Path.GetFullPath(Properties.Resources.ResourceManager.GetString("demirTiresFile"));
      txtUnipol.Text = Path.GetFullPath(Properties.Resources.ResourceManager.GetString("unipolFile"));
      txtShinService.Text = Path.GetFullPath(Properties.Resources.ResourceManager.GetString("shinServiceFile"));
      txtSaRu.Text = Path.GetFullPath(Properties.Resources.ResourceManager.GetString("saRuFile"));
    }

    private void btnMerge_Click(object sender, RoutedEventArgs e)
    {
      prbWork.IsIndeterminate = true;
      var wrk = new BackgroundWorker();
      wrk.WorkerReportsProgress = true;
      wrk.DoWork += worker_DoWork;
      wrk.ProgressChanged += worker_ProgressChanged;
      wrk.RunWorkerCompleted += worker_RunWorkerCompleted;
      wrk.RunWorkerAsync(new string[] { txtUnipol.Text, txtShinService.Text, txtSaRu.Text, txtDemirTires.Text });
    }

    private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
      //throw new NotImplementedException();
      prbWork.IsIndeterminate = false;
    }

    private void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
    {
      switch(e.ProgressPercentage)
      {
        case 1:
          lblState.Content = "Reading Unipol file...";
          break;
        case 2:
          lblState.Content = "Reading ShinService file...";
          break;
        case 3:
          lblState.Content = "Reading SaRu file...";
          break;
        case 4:
          lblState.Content = "Merging and saving result...";
          break;
        default:
          lblState.Content = "Done.";
          break;
      }
    }

    private void worker_DoWork(object sender, DoWorkEventArgs e)
    {
      var wrk = (BackgroundWorker)sender;
      wrk.ReportProgress(1);
      var args = (string[])e.Argument;
      var parameters = new Dictionary<string, object> { { "pageName", "TDSheet" }, { "id", 12 }, { "price", 13 }, { "count", 15 } };
      var uni = ExcelReader.readExcel(args[0], parameters);
      wrk.ReportProgress(2);
      parameters["id"] = 1;
      parameters["price"] = 8;
      parameters["count"] = 9;
      var shin = ExcelReader.readExcel(args[1], parameters);
      wrk.ReportProgress(3);
      parameters["pageName"] = "Легковая резина";
      parameters["id"] = 17;
      parameters["price"] = 6;
      parameters["count"] = 3;
      var sa = ExcelReader.readExcel(args[2], parameters);
      wrk.ReportProgress(4);
      try
      {
        ExcelReader.writeExcel(args[3], uni);
      }
      catch (Exception ex)
      {
        MessageBox.Show("Error saving file", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
      }
      wrk.ReportProgress(5);
      GC.Collect();
    }


    private void btnSelect_Click(object sender, RoutedEventArgs e)
    {
      var btnSrc = (Button)sender;
      var dlg = new OpenFileDialog();
      dlg.DefaultExt = "xlsx";
      dlg.Filter = "Excel workbook|*.xlsx";
      var res = dlg.ShowDialog();
      if (res.HasValue && res.Value)
      {
        switch(btnSrc.Name)
        {
          case "txtDemirTires":
            txtDemirTires.Text = dlg.FileName;
            break;
          case "txtShinService":
            txtShinService.Text = dlg.FileName;
            break;
          case "txtUnipol":
            txtUnipol.Text = dlg.FileName;
            break;
          case "txtSaRu":
            txtSaRu.Text = dlg.FileName;
            break;
        }
      }
    }
  }
}
