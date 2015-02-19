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
    public static Dictionary<string, Dictionary<string, object>> readPricelist(string file, IPricelistReader reader)
    {
      var result = new Dictionary<string, Dictionary<string, object>>();
      using (var xls = new XLWorkbook(Path.GetFullPath(file)))
      {
        var wrs = xls.Worksheet(reader.getSheetName());
        var config = reader.getParameters();
        int startRow = config["rows"].First.Value<int>("start");
        int endRow = config["rows"].First.Value<int>("end");
        for (int i = startRow; i < endRow; i++)
        {
          var prd = reader.readProduct(wrs.Row(i));
          result[prd["productId"].ToString()] = prd;
        }
      }
      return result;
    }

    public static XLWorkbook writePriceList(string file, Dictionary<string, Dictionary<string, object>> data, JToken parameters)
    {
      using (var xls = new XLWorkbook(Path.GetFullPath(file), XLEventTracking.Disabled))
      {
        return ExcelReader.writePriceList(xls, data, parameters);
      }
    }
    public static XLWorkbook writePriceList(XLWorkbook xls, Dictionary<string, Dictionary<string, object>> data, JToken parameters)
    {
      var goods = new Dictionary<string, int>();
      var wrs = xls.Worksheet(parameters.Value<string>("sheet"));
      
      var startRow = parameters["rows"].First.Value<int>("start");
      var endRow = parameters["rows"].First.Value<int>("end");
      for (var i = startRow; i <= endRow; i++)
      {
        var id = wrs.Row(i).Cell(parameters.Value<string>("productId")).Value.ToString().Trim();
        var existed = data.ContainsKey(id);
        var price = existed ? data[id]["price"] : 0;
        var quantity = existed ? data[id]["count"] : 0;
        wrs.Row(i).Cell(parameters.Value<int>("price")).Value = price;
        wrs.Row(i).Cell(parameters.Value<int>("quantity")).Value = quantity;
      }
      return xls;
    }

    public static XLWorkbook writeExcel(string file, Dictionary<string, string[]> data, Dictionary<string, object[]> parameters)
    {
      return ExcelReader.writeExcel(new XLWorkbook(Path.GetFullPath(file)), data, parameters);
    }

    public static XLWorkbook writeExcel(XLWorkbook xls, Dictionary<string, string[]> data, Dictionary<string, object[]> parameters)
    {
      foreach (var _sheet in parameters)
      {
        using (var wrs = xls.Worksheet(_sheet.Key))
        {
          var _clmnIdInd = _sheet.Value[0].ToString();
          var _clmnPriceInd = _sheet.Value[1].ToString();
          var _clmnCntInd = _sheet.Value[2].ToString();
          var rows = wrs.Rows().Where(x => !String.IsNullOrEmpty(x.Cell(_clmnIdInd).Value.ToString()));
          var goods = new Dictionary<string, int>(wrs.RowCount());
          foreach (var i in rows)
          {
            var key = i.Cell(_clmnIdInd).RichText.Text;
            if (!goods.ContainsKey(key))
              goods.Add(key, i.RowNumber());
            else
              Debug.WriteLine("Result: " + key);
          }
          foreach (var i in goods)
          {
            if (data.ContainsKey(i.Key))
            {
              var vls = data[i.Key];
              wrs.Cell(i.Value, _clmnPriceInd).Value = vls[1];
              int count = Int32.Parse(vls[2]);
              wrs.Cell(i.Value, _clmnCntInd).Value = count != 0 ? (object)count : (object)String.Empty;
            }
          }
        }
      }
      return xls;
    }
  }
}
