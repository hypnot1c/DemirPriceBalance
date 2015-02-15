using ClosedXML.Excel;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemirPriceBalance.Logic
{
  public enum TyreSeason
  {
    summer,
    winter,
    all
  }
  interface IPricelistReader
  {
    string getSheetName();
    Dictionary<string, object> readProduct(IXLRow row);
    string getProductId(IXLRow row);
    decimal getProductPrice(IXLRow row);
    int getProductCount(IXLRow row);
    string getProductManufacturer(IXLRow row);
    string getProductModel(IXLRow row);
  }
}
