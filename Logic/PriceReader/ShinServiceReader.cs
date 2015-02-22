using ClosedXML.Excel;
using DemirPriceBalance.Logic.Product;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DemirPriceBalance.Logic
{
  class ShinServiceReader : IPricelistReader
  {
    private JToken parameters;

    public ShinServiceReader(JToken parameters)
    {
      this.parameters = parameters;
    }
    public Product.Product readProduct(IXLRow row)
    {
      var result = this.getProductType(row) == ProductType.tyre ? new Tyre() : new Product.Product(); ;

      result.Id = this.getProductId(row);
      result.Price = this.getProductPrice(row);
      result.Quantity = this.getProductCount(row);

      if (String.IsNullOrEmpty(this.getProductId(row)))
        return result;

      var pType = this.getProductType(row);
      if (pType == ProductType.tyre)
        ((Tyre)result).Season = this.getTyreSeason(row);

      var str = this.parseProduct(row.Cell(4), pType, result);
      return result;
    }

    public string getSheetName()
    {
      return parameters["sheet"].ToString();
    }
    public JToken getParameters()
    {
      return this.parameters;
    }
    public string getProductId(IXLRow row)
    {
      return row.Cell(parameters["productId"].Value<int>()).Value.ToString().Trim().Split(',')[0];
    }

    public ProductType getProductType(IXLRow row)
    {
      return this.getTyreSeason(row) == TyreSeason.other ? ProductType.wheel : ProductType.tyre;
    }

    public Product.Product parseProduct(IXLCell cell, ProductType productType, Product.Product product)
    {
      switch (productType)
      {
        case ProductType.tyre:
          var _prd = (Tyre)product;
          var val = cell.Value.ToString().Trim();
          var reg = new Regex("(^|\\s)[0-9]{1,3}(/[0-9]{1,3}){0,1}(\\s|$)");
          if (reg.IsMatch(val))
          {
            var _obj = reg.Match(val).ToString().Trim().Split('/');
            _prd.ProfileWidth = _obj[0];
            _prd.ProfileHeight = _obj.Length > 1 ? _obj[1] : null;
            val = reg.Replace(val, " ");
          }
          reg = new Regex("\\s[R][0-9]{1,2}[A-Z]{0,1}\\s");
          if (reg.IsMatch(val))
          {
            _prd.Diameter = reg.Match(val).ToString().Trim().Substring(1);
            val = reg.Replace(val, " ");
          }
          reg = new Regex("\\s[0-9]{1,3}([/][0-9]{1,3}){0,1}[A-Z](\\s|$)");
          if (reg.IsMatch(val))
          {
            var str = reg.Match(val).ToString().Trim();
            _prd.WeightIndex = str.Substring(0, str.Length - 1);
            _prd.SpeedIndex = str.Last().ToString();
            if (val.IndexOf(" XL") != -1)
            {
              _prd.SpeedIndex += " XL";
              val = val.Replace(" XL", String.Empty);
            }
            val = reg.Replace(val, " ");
          }
          if (val.IndexOf(" Ш") != -1)
          {
            _prd.HasSpikes = true;
            val = val.Replace(" Ш", String.Empty);
          }
          if (val.IndexOf("RunFlat") != -1)
          {
            _prd.HasRunFlat = true;
            val = val.Replace("RunFlat", String.Empty);
          }
          _prd.Manufacturer = val.Substring(0, val.IndexOf(" "));
          _prd.Model = val.Replace(_prd.Manufacturer, String.Empty).Trim();
          break;
      }
      return product;
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
        count = value.Contains("более") ? 20 : count;
      return count;
    }

    public TyreSeason getTyreSeason(IXLRow row)
    {
      var _val = row.Cell(7).Value;
      if (_val != null)
      {
        var _strVal = _val.ToString().ToUpper();
        if (_strVal == "S")
          return TyreSeason.summer;
        if (_strVal == "W")
          return TyreSeason.winter;
      }
      return TyreSeason.other;
    }

    public string getProductManufacturer(IXLRow row)
    {
      throw new NotImplementedException();
    }

    public string getProductModel(IXLRow row)
    {
      throw new NotImplementedException();
    }

    public decimal getWheelDiameter(IXLRow row)
    {
      throw new NotImplementedException();
    }
    public decimal getWheelWidth(IXLRow row)
    {
      throw new NotImplementedException();
    }

    public uint getWheelHoles(IXLRow row)
    {
      throw new NotImplementedException();
    }
    public decimal getWheelPCD(IXLRow row)
    {
      throw new NotImplementedException();
    }
    public decimal getWheelET(IXLRow row)
    {
      throw new NotImplementedException();
    }
    public decimal getWheelDIA(IXLRow row)
    {
      throw new NotImplementedException();
    }
  }
}
