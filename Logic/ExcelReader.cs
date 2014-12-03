using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO;

using ClosedXML;
using ClosedXML.Excel;

namespace DemirPriceBalance.Logic
{
  class ExcelReader
  {
    public static Dictionary<string, string[]> readExcel(string file, Dictionary<string, object> parameters)
    {
      var _pageName = parameters["pageName"].ToString();
      var _clmnIdInd = parameters["id"].CastTo<int>();
      var _clmnPriceInd = parameters["price"].CastTo<int>();
      var _clmnCntInd = parameters["count"].CastTo<int>();
      using (var xls = new XLWorkbook(Path.GetFullPath(file)))
      {
        var wrs = xls.Worksheet(_pageName);

        var res = wrs.Rows().Where(x => !String.IsNullOrEmpty(x.Cell(_clmnIdInd).Value.CastTo<String>()) && x.Cell(_clmnIdInd).Value.CastTo<String>() != "Код производителя");
        var goods = new Dictionary<string, string[]>(res.Count());
        foreach (var i in res)
        {
          var key = i.Cell(_clmnIdInd).RichText.Text;
          if (!goods.ContainsKey(key))
            goods.Add(key, new string[] { i.RowNumber().ToString(), i.Cell(_clmnPriceInd).RichText.Text, i.Cell(_clmnCntInd).RichText.Text });
          else
            Debug.WriteLine(key);
        }
        return goods;
      }
    }

    public static void writeExcel(string file, Dictionary<string, string[]> data)
    {
      using (var xls = new XLWorkbook(Path.GetFullPath(file)))
      {
        var sheets = new Dictionary<string, Dictionary<string, object>> { {"Шины", new Dictionary<string, object> { { "price", (object)"X" } } } };
        using (var wrs = xls.Worksheet("Шины"))
        {
          var rows = wrs.Rows().Where(x => !String.IsNullOrEmpty(x.Cell("A").Value.ToString()));
          var goods = new Dictionary<string, int>(wrs.RowCount());
          foreach (var i in rows)
          {
            var key = i.Cell("A").RichText.Text;
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
              wrs.Cell(i.Value, "X").Value = vls[1];
              int count = 0;
              var isNum = Int32.TryParse(vls[2], out count);
              wrs.Cell(i.Value, "W").Value = isNum ? (object)count : (object)String.Empty;
            }
          }
        }
        xls.SaveAs(@"D:\Documents\GitHub\DemirPriceBalance\docs\DEMIR шины и диски 20.10.20141.xlsx");
      }
    }
  }
}
