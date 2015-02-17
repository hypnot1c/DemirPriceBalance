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
    all,
    other
  }

  public enum ProductType
  {
    wheel,
    tyre,
    other
  }
  interface IPricelistReader
  {
    string getSheetName();
    JToken getParameters();
    Dictionary<string, object> readProduct(IXLRow row);
    string getProductId(IXLRow row);
    ProductType getProductType(IXLRow row);
    Dictionary<string, object> parseProduct(IXLCell cell, ProductType productType);
    decimal getProductPrice(IXLRow row);
    int getProductCount(IXLRow row);
    string getProductManufacturer(IXLRow row);
    string getProductModel(IXLRow row);
    TyreSeason getTyreSeason(IXLRow row);
  }
}
