using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO;

using ClosedXML;
using ClosedXML.Excel;
using DemirPriceBalance.Logic;
using Newtonsoft.Json.Linq;

namespace DemirPriceBalance.Logic
{
  class ExcelReader
  {
    public static Dictionary<string, Product.Product> readPricelist(string file, IPricelistReader reader)
    {
      var result = new Dictionary<string, Product.Product>();
      using (var xls = new XLWorkbook(Path.GetFullPath(file)))
      {
        var wrs = xls.Worksheet(reader.getSheetName());
        var config = reader.getParameters();
        int startRow = config["rows"].First.Value<int>("start");
        int endRow = config["rows"].First.Value<int>("end");
        for (int i = startRow; i < endRow; i++)
        {
          var prd = reader.readProduct(wrs.Row(i));
          if(!String.IsNullOrEmpty(prd.Id))
            result[prd.Id] = prd;
        }
      }
      return result;
    }

    public static XLWorkbook writePriceList(string file, Dictionary<string, Product.Product> data, JToken parameters)
    {
      using (var xls = new XLWorkbook(Path.GetFullPath(file), XLEventTracking.Disabled))
      {
        return ExcelReader.writePriceList(xls, data, parameters);
      }
    }
    public static XLWorkbook writePriceList(XLWorkbook xls, Dictionary<string, Product.Product> data, JToken parameters)
    {
      var goods = new Dictionary<string, int>();
      var wrs = xls.Worksheet(parameters.Value<string>("sheet"));
      
      var startRow = parameters["rows"].First.Value<int>("start");
      var endRow = parameters["rows"].First.Value<int>("end");
      var clmnPrice = parameters.Value<int>("price");
      var clmnQnty = parameters.Value<int>("quantity");
      var pId = parameters.Value<string>("productId");
      var prds = wrs.Rows(startRow, endRow).ToDictionary(x => x.Cell(pId).Value.ToString().Trim(), x => x.RowNumber());
      wrs.Range(startRow, clmnPrice, endRow, clmnPrice).SetValue<decimal>(0);
      wrs.Range(startRow, clmnQnty, endRow, clmnQnty).SetValue<int>(0);
      var lastRowIndex = wrs.LastRowUsed().RowNumber() + 1;
      var style = wrs.Row(lastRowIndex - 1).Style;
      foreach (var prd in data)
      {
        if (prds.ContainsKey(prd.Key))
        {
          wrs.Cell(prds[prd.Key], clmnPrice).Value = prd.Value.Price;
          wrs.Cell(prds[prd.Key], clmnQnty).Value = prd.Value.Quantity;
        }
        else
        {
          wrs.Row(lastRowIndex).Style = style;
          var cell = wrs.Cell(lastRowIndex++, 1);
          var row = new List<object[]>() { prd.Value.ToExcelRow() };
          cell.InsertData(row.AsEnumerable());
          
        }
      }
      return xls;
    }
  }
}
