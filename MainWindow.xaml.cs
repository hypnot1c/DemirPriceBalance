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
      txtDemirTires.Text = Path.GetFullPath(Properties.Resources.ResourceManager.GetString("demirTiresInFile"));
      txtDemirTiresSrc.Text = Path.GetFullPath(Properties.Resources.ResourceManager.GetString("demirTiresFile"));
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
      wrk.RunWorkerAsync(new string[] { txtUnipol.Text, txtShinService.Text, txtSaRu.Text, txtDemirTiresSrc.Text, txtDemirTires.Text });
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
      var parameters = new Dictionary<string, object> { { "pageName", "TDSheet" }, { "id", 12 }, { "price", 14 }, { "count", 15 } };
      var uni = ExcelReader.readExcel(args[0], parameters);
      wrk.ReportProgress(2);
      parameters["id"] = 1;
      parameters["price"] = 8;
      parameters["count"] = 9;
      var shin = ExcelReader.readExcel(args[1], parameters);
      wrk.ReportProgress(3);
      parameters["pageName"] = "Диски";
      parameters["id"] = 1;
      parameters["price"] = 6;
      parameters["count"] = 3;
      var sa = ExcelReader.readExcel(args[2], parameters);
      wrk.ReportProgress(4);
      try
      {
        var pars = new Dictionary<string, object[]> { { "Шины", new object[] { 1, 24, 23 } } };
        using (var xls = ExcelReader.writeExcel(args[3], uni, pars))
        {
          pars = new Dictionary<string, object[]> { { "Шины", new object[] { 1, 26, 25 } } };
          var xls1 = ExcelReader.writeExcel(xls, shin, pars);
          pars = new Dictionary<string, object[]> { { "Диски реплика", new object[] { 1, 20, 19 } } };
          xls1 = ExcelReader.writeExcel(xls, sa, pars);
          xls1.SaveAs(args[4]);
        }
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
    private void workerSQL_DoWork(object sender, DoWorkEventArgs e)
    {
      var parameters = new Dictionary<string, object> { { "pageName", "Шины" }, { "id", 1 }, { "price", 16 }, { "count", 15 } };
      var uni = ExcelReader.readExcel((string)e.Argument, parameters);
      parameters = new Dictionary<string, object> { { "pageName", "Диски реплика" }, { "id", 1 }, { "price", 13 }, { "count", 12 } };
      var sa = ExcelReader.readExcel((string)e.Argument, parameters);
      var res = uni.Where(x => Int32.Parse(x.Value[2]) > 3).Select(x => String.Concat("UPDATE `oc_product` SET `quantity` = ", x.Value[2], ", `price` = ", x.Value[1], " WHERE `sku` = \"", x.Key, "\""));
      var res2 = sa.Where(x => Int32.Parse(x.Value[2]) > 3).Select(x => String.Concat("UPDATE `oc_product` SET `quantity` = ", x.Value[2], ", `price` = ", x.Value[1], " WHERE `sku` = \"", x.Key, "\""));
      
      File.WriteAllLines(@"C:\Users\hypnotic\Documents\GitHub\DemirPriceBalance\DemirPriceBalance\docs\query.sql", res.Concat(res2));
    }
    private void button_Click(object sender, RoutedEventArgs e)
    {
      prbWork.IsIndeterminate = true;
      var wrk = new BackgroundWorker();
      wrk.DoWork += workerSQL_DoWork;
      wrk.RunWorkerCompleted += worker_RunWorkerCompleted;
      wrk.RunWorkerAsync(txtDemirTires.Text);
    }
  }
}
