using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemirPriceBalance.Logic.PriceReader
{
  class DemirWheelsReader : IPricelistReader
  {
    public string getSheetName()
    {
      throw new NotImplementedException();
    }

    public Newtonsoft.Json.Linq.JToken getParameters()
    {
      throw new NotImplementedException();
    }

    public Product.Product readProduct(ClosedXML.Excel.IXLRow row)
    {
      throw new NotImplementedException();
    }

    public string getProductId(ClosedXML.Excel.IXLRow row)
    {
      throw new NotImplementedException();
    }

    public ProductType getProductType(ClosedXML.Excel.IXLRow row)
    {
      throw new NotImplementedException();
    }

    public Product.Product parseProduct(ClosedXML.Excel.IXLCell cell, ProductType productType, Product.Product product)
    {
      throw new NotImplementedException();
    }

    public decimal getProductPrice(ClosedXML.Excel.IXLRow row)
    {
      throw new NotImplementedException();
    }

    public uint getProductCount(ClosedXML.Excel.IXLRow row)
    {
      throw new NotImplementedException();
    }

    public string getProductManufacturer(ClosedXML.Excel.IXLRow row)
    {
      throw new NotImplementedException();
    }

    public string getProductModel(ClosedXML.Excel.IXLRow row)
    {
      throw new NotImplementedException();
    }

    public decimal getWheelDiameter(ClosedXML.Excel.IXLRow row)
    {
      throw new NotImplementedException();
    }

    public decimal getWheelWidth(ClosedXML.Excel.IXLRow row)
    {
      throw new NotImplementedException();
    }

    public uint getWheelHoles(ClosedXML.Excel.IXLRow row)
    {
      throw new NotImplementedException();
    }

    public decimal getWheelPCD(ClosedXML.Excel.IXLRow row)
    {
      throw new NotImplementedException();
    }

    public decimal getWheelET(ClosedXML.Excel.IXLRow row)
    {
      throw new NotImplementedException();
    }

    public decimal getWheelDIA(ClosedXML.Excel.IXLRow row)
    {
      throw new NotImplementedException();
    }

    public Product.TyreSeason getTyreSeason(ClosedXML.Excel.IXLRow row)
    {
      throw new NotImplementedException();
    }
  }
}
