using System;
using System.Diagnostics;
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
using Microsoft.Win32;

using ClosedXML;
using ClosedXML.Excel;

using DemirPriceBalance.Logic;
using DemirPriceBalance.Logic.PriceReader;
using DemirPriceBalance.Logic.Product;
using Newtonsoft.Json.Linq;

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
      Config.LoadConfig();
      txtDemirTires.Text = Path.GetFullPath(Config.data["outputFile"]["path"].ToString());
      var inpFiles = (JArray)Config.data["inputFiles"];
      for (var i = 0; i < inpFiles.Count; i++)
      {
        var txt = new TextBox();
        txt.Width = 240;
        txt.Height = 24;
        txt.Margin = new Thickness(10, 10 + (i * 30), 0, 0);
        txt.VerticalAlignment = System.Windows.VerticalAlignment.Top;
        txt.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
        txt.Visibility = System.Windows.Visibility.Visible;
        txt.Text = Path.GetFullPath(inpFiles[i]["srcFileConfig"]["path"].ToString());
        grdInputFiles.Children.Add(txt);
        var btn = new Button();
        btn.Width = 24;
        btn.Height = 24;
        btn.VerticalAlignment = System.Windows.VerticalAlignment.Top;
        btn.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
        btn.Margin = new Thickness(255, 10 + (i * 30), 0, 0);
        var sp = new StackPanel();
        var img = new Image();
        img.Source = (ImageSource)FindResource("upload");
        sp.Children.Add(img);
        btn.Content = sp;
        grdInputFiles.Children.Add(btn);
      }
    }

    private void btnMerge_Click(object sender, RoutedEventArgs e)
    {
      prbWork.IsIndeterminate = true;
      var wrk = new BackgroundWorker();
      wrk.WorkerReportsProgress = true;
      wrk.DoWork += worker_DoWork;
      wrk.ProgressChanged += worker_ProgressChanged;
      wrk.RunWorkerCompleted += worker_RunWorkerCompleted;
      var inpts = new List<string>();
      foreach (var cnt in grdInputFiles.Children)
        if (cnt.GetType().Name == "TextBox")
          inpts.Add(((TextBox)cnt).Text);
      wrk.RunWorkerAsync(inpts);
    }

    private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
      prbWork.IsIndeterminate = false;
    }

    private void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
    {
      if (e.ProgressPercentage == 100)
      {
        lblState.Content = "Done";
      }
      else
      {
        var provider = (string)e.UserState;
        lblState.Content = String.Concat("Reading ", provider, "...");
      }
    }

    private void worker_DoWork(object sender, DoWorkEventArgs e)
    {
      var wrk = (BackgroundWorker)sender;
      var args = (List<string>)e.Argument;
      var fileData = new List<Dictionary<string, Product>>(args.Count);
      for (var i = 0; i < args.Count; i++)
      {
        var provider = Config.data["inputFiles"][i].Value<string>("provider");
        wrk.ReportProgress(i, provider);
        fileData.Add(ExcelReader.readPricelist(args[i], PriceReaderFactory.getPriceReader(provider, Config.data["inputFiles"][i]["srcFileConfig"])));
      }

      var xls = new XLWorkbook(Config.data["outputFile"].Value<string>("path"), XLEventTracking.Disabled);
      for(var i = 0; i < fileData.Count; i++) 
        xls = ExcelReader.writePriceList(xls, fileData[i], Config.data["inputFiles"][i]["trgFileConfig"]);
      
      xls.Save();
      wrk.ReportProgress(100);
    }


    private void btnSelect_Click(object sender, RoutedEventArgs e)
    {
      var btnSrc = (Button)sender;
      var dlg = new OpenFileDialog();
      dlg.DefaultExt = "xlsx";
      dlg.Filter = "Excel workbook|*.xlsx;*.xlsm";
      var res = dlg.ShowDialog();
      if (res.HasValue && res.Value)
      {
        switch(btnSrc.Name)
        {
          case "btnDemirTiresSrc":
            //txtDemirTiresSrc.Text = dlg.FileName;
            break;
          case "btnShinService":
            //txtShinService.Text = dlg.FileName;
            break;
          case "btnUnipol":
            //txtUnipol.Text = dlg.FileName;
            break;
          case "btnSaRu":
            //txtSaRu.Text = dlg.FileName;
            break;
          case "btnSQLfile":
            txtSQLfile.Text = dlg.FileName;
            break;
        }
      }
    }
    private void workerSQL_DoWork(object sender, DoWorkEventArgs e)
    {
      //var parameters = new Dictionary<string, object> { { "pageName", "Шины" }, { "id", 1 }, { "price", 16 }, { "count", 15 } };
      //var uni = ExcelReader.readExcel((string)e.Argument, parameters);
      //parameters = new Dictionary<string, object> { { "pageName", "Диски реплика" }, { "id", 1 }, { "price", 13 }, { "count", 12 } };
      //var sa = ExcelReader.readExcel((string)e.Argument, parameters);
      //parameters = new Dictionary<string, object> { { "pageName", "Диски тюнинг" }, { "id", 1 }, { "price", 13 }, { "count", 12 } };
      //var sa1 = ExcelReader.readExcel((string)e.Argument, parameters);
      //var res = uni.Select(x => String.Concat("INSERT INTO `tmp_Import` (`id`, `price`, `count`) VALUES (\"", x.Key, "\", ", x.Value[1], ", ", x.Value[2], ");"));
      //var res2 = sa.Select(x => String.Concat("INSERT INTO `tmp_Import` (`id`, `price`, `count`) VALUES (\"", x.Key, "\", ", x.Value[1], ", ", x.Value[2], ");"));
      //var res3 = sa1.Select(x => String.Concat("INSERT INTO `tmp_Import` (`id`, `price`, `count`) VALUES (\"", x.Key, "\", ", x.Value[1], ", ", x.Value[2], ");"));

      //File.WriteAllLines(@"C:\Users\hypnotic\Documents\GitHub\DemirPriceBalance\DemirPriceBalance\docs\query.sql", res.Concat(res2).Concat(res3));
    }
    private void button_Click(object sender, RoutedEventArgs e)
    {
      prbWork.IsIndeterminate = true;
      var wrk = new BackgroundWorker();
      wrk.DoWork += workerSQL_DoWork;
      wrk.RunWorkerCompleted += worker_RunWorkerCompleted;
      wrk.RunWorkerAsync(txtSQLfile.Text);
    }

    private void btnGenYML_Click(object sender, RoutedEventArgs e)
    {
      //var offers = new List<Offer>();
      //using (var xls = new XLWorkbook(txtDemirTires.Text))
      //{
      //  var sheet = xls.Worksheet("Шины");
      //  for(var i = 4; i <= sheet.RowCount(); i++)
      //  {
      //    var count = sheet.Cell(i, 15).ValueCached == null ? sheet.Cell(i, 15).Value.ToString() : sheet.Cell(i, 15).ValueCached;

      //    var price = sheet.Cell(i, 16).ValueCached == null ? sheet.Cell(i, 16).Value.ToString() : sheet.Cell(i, 16).ValueCached;
      //    offers.Add(new Offer(
      //                   OfferType.vendorModel,
      //                   ExcelReader.GetProductCount(count) > 3 ? true : false,
      //                   String.Empty,
      //                   price,
      //                   "RUR",
      //                   1,
      //                   String.Empty,
      //                   String.Empty,
      //                   sheet.Cell(i, 2).Value.ToString(),
      //                   sheet.Cell(i, 3).Value.ToString()
      //                   ));
      //  }
      //}
      //GC.Collect();
      //var yml = new YML(@"C:\Users\hypnotic\Documents\GitHub\DemirPriceBalance\DemirPriceBalance\docs\demirshinidiski.yml", @"C:\Users\hypnotic\Documents\GitHub\DemirPriceBalance\DemirPriceBalance\docs\demirshinidiski_res.yml");
      //yml.writeOffers(offers.ToArray());
    }
  }
}
