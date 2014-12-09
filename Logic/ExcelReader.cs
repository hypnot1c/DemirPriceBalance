using System;
using System.Collections;
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
    public static EventHandler readDone;
    public static EventHandler writeDone;

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

    private static int _documentCount = 0;
    public static int DocumentCount
    {
      get { return _documentCount; }
      set
      {
        _documentCount = value;
      }
    }

    public static Queue<PriceList> DocData = new Queue<PriceList>();

    public static void readExcel(object data)
    {
      var _data = (Dictionary<string, object>)data;
      var file = _data["file"].ToString();
      var parameters = (Dictionary<string, object>)_data["parameters"];
      var supplierName = _data["supplierName"].ToString();
      var _pageName = parameters["pageName"].ToString();
      var _clmnIdInd = parameters["id"].CastTo<int>();
      var _clmnPriceInd = parameters["price"].CastTo<int>();
      var _clmnCntInd = parameters["count"].CastTo<int>();
      using (var xls = new XLWorkbook(Path.GetFullPath(file)))
      {
        var wrs = xls.Worksheet(_pageName);

        var res = wrs.Rows().Where(x => !String.IsNullOrEmpty(x.Cell(_clmnIdInd).Value.CastTo<String>()) && x.Cell(_clmnIdInd).Value.CastTo<String>() != "Код производителя");
        var goods = new List<PriceList.Stuff>(res.Count());
        foreach (var i in res)
        {
          var key = i.Cell(_clmnIdInd).RichText.Text;
          var _count = ExcelReader.GetProductCount(i.Cell(_clmnCntInd).RichText.Text);
          goods.Add(new PriceList.Stuff(key, _count == 0 ? String.Empty : _count.ToString(), i.Cell(_clmnPriceInd).RichText.Text));
        }

        lock (DocData)
        {
          DocData.Enqueue(new PriceList(supplierName, goods));
          _documentCount--;
          if (readDone != null)
            readDone(new object(), EventArgs.Empty);
        }
      }
    }

    public static XLWorkbook writeExcel(object data)
    {
      var _pars = (Dictionary<string, object>)data;
      var xls = _pars["file"].GetType().ToString() == "string" ? new XLWorkbook(Path.GetFullPath(_pars["file"].ToString()).ToString()) : (XLWorkbook)_pars["file"];
      var _data = (PriceList)_pars["data"];
      var parameters = (Dictionary<string, object[]>)_pars["parameters"];

      foreach (var _sheet in parameters)
      {
        using (var wrs = xls.Worksheet(_sheet.Key))
        {
          var _clmnIdInd = _sheet.Value[0].ToString();
          var _clmnPriceInd = _sheet.Value[1].ToString();
          var _clmnCntInd = _sheet.Value[2].ToString();
          var goods = new Dictionary<string, int>(wrs.RowCount());
          for (var _i = 1; _i <= wrs.RowCount(); _i++)
          {
            if (String.IsNullOrEmpty(wrs.Cell(_i, _clmnIdInd).Value.ToString())) continue;

            var key = wrs.Cell(_i, _clmnIdInd).RichText.Text;
            if (!goods.ContainsKey(key))
              goods.Add(key, _i);
            else
              Debug.WriteLine("Result: " + key);
          }

          for (var _i = 0; _i < _data.Goods.Count; _i++)
          {
            var vls = _data.Goods[_i];
            if (goods.ContainsKey(vls.id))
            {
              wrs.Cell(goods[vls.id], _clmnPriceInd).Value = vls.price;
              int count = 0;
              var isNum = Int32.TryParse(vls.count, out count);
              wrs.Cell(goods[vls.id], _clmnCntInd).Value = isNum ? (object)count : (object)String.Empty;
            }
          }
        }
      }
      return xls;
    }
  }
}
