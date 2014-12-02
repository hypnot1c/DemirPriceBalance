using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

using ClosedXML;
using ClosedXML.Excel;

namespace DemirPriceBalance.Logic
{
  class ExcelReader
  {
    public static void readExcel()
    {
      using (var xls = new XLWorkbook(@"D:\Documents\Demir\Юнипол.xlsx"))
      {
        var wrs = xls.Worksheet("TDSheet");
        var rowCnt = wrs.RowCount();
        var res = wrs.Rows().Where(x => !String.IsNullOrEmpty(x.Cell(12).Value.CastTo<String>()) && x.Cell(12).Value.CastTo<String>() != "Код производителя");
        var goods = new Dictionary<string, string[]>(res.Count());
        try
        {
          foreach (var i in res)
          {
            var key = i.Cell(12).RichText.Text;
            if (!goods.ContainsKey(key))
              goods.Add(key, new string[] { i.RowNumber().ToString(), i.Cell(13).RichText.Text, i.Cell(15).RichText.Text });
            else
              Debug.WriteLine(key);
          }
        }
        catch(Exception ex)
        {
          ex.ToString();
        }
        //res.Count();

        using(var xls2 = new XLWorkbook(@"D:\Documents\Demir\DEMIR шины и диски 20.10.2014.xlsx"))
        {
          wrs = xls2.Worksheet("Шины");
          var t = wrs.Row(4).Cell(16).Value;
          var inr = wrs.Rows().Select(x => x.Cell(1).Value.CastTo<String>()).Intersect(goods.Keys);
        }
        using (var xls3 = new XLWorkbook(@"D:\Documents\Demir\v8_2BA4_378.xlsx"))
        {
          wrs = xls3.Worksheet("TDSheet");
          var inr = wrs.Rows().Select(x => x.Cell(1).Value.CastTo<String>()).Intersect(goods.Keys);
        }
      }
    }
  }
}
