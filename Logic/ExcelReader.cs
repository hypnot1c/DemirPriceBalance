using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ClosedXML;
using ClosedXML.Excel;

namespace DemirPriceBalance.Logic
{
  class ExcelReader
  {
    public static void readExcel()
    {
      var xls = new XLWorkbook(@"C:\Users\hypnotic\Documents\DemirGroup\price\Юнипол.xlsx");
      var wrs = xls.Worksheet("TDSheet");
      var rowCnt = wrs.RowCount();
      var res = wrs.Rows().Where(x => !String.IsNullOrEmpty(x.Cell(12).Value.CastTo<String>())).Select(x => new string[] { x.Cell(12).Value.CastTo<String>(), x.Cell(13).Value.CastTo<String>(), x.Cell(15).Value.CastTo<String>() });
      res.Count();
    }
  }
}
