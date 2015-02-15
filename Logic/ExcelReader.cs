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

namespace DemirPriceBalance.Logic
{
  class ExcelReader
  {
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

    public static List<Dictionary<string, object>> readPricelist(string file, IPricelistReader reader)
    {
      var result = new List<Dictionary<string, object>>();
      using (var xls = new XLWorkbook(Path.GetFullPath(file)))
      {
        var wrs = xls.Worksheet(reader.getSheetName());
        for (int i = 4074; i < 8255; i++)
        {
          result.Add(reader.readProduct(wrs.Row(i)));
        }
      }
      return result;
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
