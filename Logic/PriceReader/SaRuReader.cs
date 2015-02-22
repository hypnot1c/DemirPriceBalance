using ClosedXML.Excel;
using DemirPriceBalance.Logic.Product;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DemirPriceBalance.Logic.PriceReader
{
  class SaRuReader : IPricelistReader
  {
    private JToken parameters;

    public SaRuReader(JToken parameters)
    {
      this.parameters = parameters;
    }
    public string getSheetName()
    {
      return parameters["sheet"].ToString();
    }

    public JToken getParameters()
    {
      return this.parameters;
    }

    public Product.Product readProduct(IXLRow row)
    {
      var result = new Wheel();
      result.Id = this.getProductId(row);
      result.Price = this.getProductPrice(row);
      result.Quantity = this.getProductCount(row);

      if (String.IsNullOrEmpty(this.getProductId(row)))
        return result;

      result.Manufacturer = this.getProductManufacturer(row);
      result.Model = this.getProductModel(row);
      result.Diameter = this.getWheelDiameter(row);
      result.Width = this.getWheelWidth(row);
      result.Holes = this.getWheelHoles(row);
      result.PCD = this.getWheelPCD(row);
      result.ET = this.getWheelET(row);
      result.DIA = this.getWheelDIA(row);
      return result;
    }

    public string getProductId(IXLRow row)
    {
      return row.Cell(parameters["productId"].Value<int>()).Value.ToString().Trim().Split(',')[0];
    }

    public ProductType getProductType(IXLRow row)
    {
      throw new NotImplementedException();
    }

    public Product.Product parseProduct(IXLCell cell, ProductType productType, Product.Product product)
    {
      throw new NotImplementedException();
    }

    public decimal getProductPrice(IXLRow row)
    {
      var value = row.Cell(parameters["price"].Value<int>()).Value;
      decimal price = 0;
      Decimal.TryParse(value.ToString().Trim(), out price);
      return price;
    }

    public uint getProductCount(IXLRow row)
    {
      var value = row.Cell(parameters["quantity"].Value<int>()).Value.ToString().ToLower().Trim();
      uint count = 0;
      if (!UInt32.TryParse(value, out count))
        count = value.Contains("да") ? 20 : count;
      return count;
    }

    public string getProductManufacturer(IXLRow row)
    {
      return row.Cell(10).Value.ToString().Trim();
    }

    public string getProductModel(IXLRow row)
    {
      return row.Cell(11).Value.ToString().Trim();
    }
    public decimal getWheelDiameter(IXLRow row)
    {
      var val = row.Cell(12).Value.ToString();
      decimal diam = 0;
      Decimal.TryParse(val, out diam);
      return diam;
    }
    public decimal getWheelWidth(IXLRow row)
    {
      var val = row.Cell(12).Value.ToString();
      decimal diam = 0;
      Decimal.TryParse(val, NumberStyles.Any, CultureInfo.InvariantCulture, out diam);
      return diam;
    }
    public uint getWheelHoles(IXLRow row)
    {
      var val = row.Cell(14).Value.ToString();
      uint diam = 0;
      UInt32.TryParse(val, out diam);
      return diam;
    }
    public decimal getWheelPCD(IXLRow row)
    {
      var val = row.Cell(15).Value.ToString();
      decimal diam = 0;
      Decimal.TryParse(val, out diam);
      return diam;
    }
    public decimal getWheelET(IXLRow row)
    {
      var val = row.Cell(17).Value.ToString();
      decimal diam = 0;
      Decimal.TryParse(val, out diam);
      return diam;
    }
    public decimal getWheelDIA(IXLRow row)
    {
      var val = row.Cell(18).Value.ToString();
      decimal diam = 0;
      Decimal.TryParse(val, out diam);
      return diam;
    }

    public Product.TyreSeason getTyreSeason(IXLRow row)
    {
      throw new NotImplementedException();
    }
  }
}
