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
using System.Threading;

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
      var uniThread = new Thread(ExcelReader.readExcel);
      var shinThread = new Thread(ExcelReader.readExcel);
      var saThread = new Thread(ExcelReader.readExcel);
      ExcelReader.readDone += worker_RunWorkerCompleted;
      var _pars1 = new Dictionary<string, object> {
        { "file", txtUnipol.Text },
        { "supplierName", "uni" },
        { "parameters", new Dictionary<string, object> { { "pageName", "TDSheet" }, { "id", 12 }, { "price", 14 }, { "count", 15 } } },
      };
      var _pars2 = new Dictionary<string, object> {
        { "file", txtShinService.Text },
        { "supplierName", "shin" },
        { "parameters", new Dictionary<string, object> { { "pageName", "TDSheet" }, { "id", 1 }, { "price", 8 }, { "count", 9 } } }
      };
      var _pars3 = new Dictionary<string, object> {
        { "file", txtSaRu.Text },
        { "supplierName", "sa" },
        { "parameters", new Dictionary<string, object> { { "pageName", "Диски" }, { "id", 1 }, { "price", 6 }, { "count", 3 } } }
      };
      ExcelReader.DocumentCount = 3;
      lblState.Content = "Reading files. 3 left...";
      uniThread.Start(_pars1);
      shinThread.Start(_pars2);
      saThread.Start(_pars3);
    }

    private Dictionary<string, object[]> GetWriteParams(string supplierName)
    {
      switch (supplierName)
      {
        case "uni":
          return new Dictionary<string, object[]> { { "Шины", new object[] { 1, 24, 23 } } };
        case "shin":
          return new Dictionary<string, object[]> { { "Шины", new object[] { 1, 26, 25 } } };
        case "sa":
          return new Dictionary<string, object[]> { { "Диски реплика", new object[] { 1, 20, 19 } } };
        default:
          return new Dictionary<string, object[]>();
      }
    }

    public void worker_RunWorkerCompleted(object sender, EventArgs e)
    {
      if (ExcelReader.DocumentCount == 0)
      {
        //File.Copy(Path.GetFullPath("../../docs/DEMIR шины и диски 20.10.2014.xlsx"), Path.GetFullPath("../../docs/DEMIR_Tires_and_Disks.xlsx"), true);
        //var xls = new ClosedXML.Excel.XLWorkbook(Path.GetFullPath("../../docs/DEMIR_Tires_and_Disks.xlsx"));
        //while(ExcelReader.DocData.Count > 0) {
        //  var supp = ExcelReader.DocData.Dequeue();
        //  var pars = this.GetWriteParams(supp.SupplierName);
        //  var _pars = new Dictionary<string, object> { { "file", xls}, {  "data", supp }, { "parameters", pars} };
        //  xls = ExcelReader.writeExcel(_pars);
        //}
        //xls.Save();
        lblState.Dispatcher.Invoke(delegate() { lblState.Content = "Done."; });
        GC.Collect();
      }
      else
      {
        lblState.Dispatcher.Invoke(delegate() { lblState.Content = "Reading files. " + ExcelReader.DocumentCount.ToString() + " left..."; });
      }
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
