using ClosedXML.Excel;
using DemirPriceBalance.Logic.Product;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemirPriceBalance.Logic
{
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
    Product.Product readProduct(IXLRow row);
    string getProductId(IXLRow row);
    ProductType getProductType(IXLRow row);
    Product.Product parseProduct(IXLCell cell, ProductType productType, Product.Product product);
    decimal getProductPrice(IXLRow row);
    uint getProductCount(IXLRow row);
    string getProductManufacturer(IXLRow row);
    string getProductModel(IXLRow row);
    decimal getWheelDiameter(IXLRow row);
    decimal getWheelWidth(IXLRow row);
    uint getWheelHoles(IXLRow row);
    decimal getWheelPCD(IXLRow row);
    decimal getWheelET(IXLRow row);
    decimal getWheelDIA(IXLRow row);
    TyreSeason getTyreSeason(IXLRow row);
  }
}
