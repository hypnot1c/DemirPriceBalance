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
          if (!String.IsNullOrEmpty(prd.Id))
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
      var lastRowIndex = wrs.LastRowUsed().RowNumber();
      var prds = wrs.Rows(startRow, lastRowIndex++).ToDictionary(x => x.Cell(pId).Value.ToString().Trim(), x => x.RowNumber());
      wrs.Range(startRow, clmnPrice, endRow, clmnPrice).SetValue<decimal>(0);
      wrs.Range(startRow, clmnQnty, endRow, clmnQnty).SetValue<int>(0);
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
          var row = new List<object[]>() { prd.Value.ToExcelRow(clmnQnty, clmnPrice) };
          cell.InsertData(row.AsEnumerable());
        }
      }
      return xls;
    }
    public static int GetProductCount(string sourceValue)
    {
      var _value = sourceValue.ToLower().Trim();
      int _result = 0;
      var isNum = Int32.TryParse(_value, out _result);
      if (!isNum)
      {
        if (_value == "да" || _value.StartsWith("более"))
          _result = 20;
      }
      return _result;
    }
    public static Dictionary<string, string[]> readExcel(string file, Dictionary<string, object> parameters)
    {
      var _pageName = parameters["pageName"].ToString();
      var _clmnIdInd = parameters["id"].CastTo<int>();
      var _clmnPriceInd = parameters["price"].CastTo<int>();
      var _clmnCntInd = parameters["count"].CastTo<int>();
      var result = new List<Dictionary<string, object>>();
      using (var xls = new XLWorkbook(Path.GetFullPath(file)))
      {
        var wrs = xls.Worksheet(_pageName);

        var res = wrs.Rows().Where(x => !String.IsNullOrEmpty(x.Cell(_clmnIdInd).Value.CastTo<String>()) && x.Cell(_clmnIdInd).Value.CastTo<String>() != "Код производителя");
        var goods = new Dictionary<string, string[]>(res.Count());
        foreach (var i in res)
        {
          var keys = i.Cell(_clmnIdInd).RichText.Text.Split(',');
          foreach (var key in keys)
          {
            if (!goods.ContainsKey(key))
            {
              var _value = i.Cell(_clmnCntInd).ValueCached == null ? i.Cell(_clmnCntInd).Value.ToString() : i.Cell(_clmnCntInd).ValueCached;
              var _price = i.Cell(_clmnPriceInd).ValueCached == null ? i.Cell(_clmnPriceInd).Value.ToString() : i.Cell(_clmnPriceInd).ValueCached;
              var _count = ExcelReader.GetProductCount(_value);
              var _price1 = ExcelReader.GetProductCount(_price);
              goods.Add(key, new string[] { i.RowNumber().ToString(), _price1.ToString(), _count.ToString() });
            }
            else
              Debug.WriteLine(key);
          }
        }
        return goods;
      }
    }
  }
}
