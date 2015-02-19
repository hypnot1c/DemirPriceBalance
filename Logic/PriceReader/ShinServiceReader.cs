using ClosedXML.Excel;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DemirPriceBalance.Logic
{
  class ShinServiceReader: IPricelistReader
  {
    private JToken parameters;

    public ShinServiceReader(JToken parameters)
    {
      this.parameters = parameters;
    }
    public Dictionary<string, object> readProduct(IXLRow row)
    {
      var result = new Dictionary<string, object>();
      result["productId"] = this.getProductId(row);
      result["price"] = this.getProductPrice(row);
      result["count"] = this.getProductCount(row);
      
      var pType = this.getProductType(row);
      if(pType == ProductType.tyre)
        result["season"] = this.getTyreSeason(row);
      
      if(String.IsNullOrEmpty(this.getProductId(row))) 
          return result;

      var str = this.parseProduct(row.Cell(4), pType);

      return result.Union(str).ToDictionary(x => x.Key, x => x.Value);
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
      return row.Cell(parameters["productId"].Value<int>()).Value.ToString().Trim();
    }

    public ProductType getProductType(IXLRow row)
    {
      return this.getTyreSeason(row) == TyreSeason.other ? ProductType.wheel : ProductType.tyre;
    }

    public Dictionary<string, object> parseProduct(IXLCell cell, ProductType productType)
    {
      var result = new Dictionary<string, object>();
      result["productType"] = productType;
      switch (productType)
      {
        case ProductType.tyre:
          var val = cell.Value.ToString().Trim();
          var sb = new StringBuilder();
          for (int i = 0, j = 0; i < val.Length; i++)
          {
            if (val[i] != ' ')
            {
              sb.Append(val[i]);
            }
            else
            {
              switch (j++)
              {
                case 0:
                  result["manufacturer"] = sb.ToString();
                  break;
                case 1:
                  var _obj = sb.ToString().Split('/');
                  result["width"] = _obj[0];
                  result["profile"] = _obj.Length > 1 ? _obj[1] : null;
                  break;
                case 2:
                  result["diameter"] = sb.ToString();
                  break;
                case 3:
                  var reg = new Regex("^([0-9]{1,3}[/]{0,1}){0,1}[0-9]{1,3}[A-Z]$");
                  var modelPart = new StringBuilder();
                  for (i++; i < val.Length; i++) {
                    if (val[i] == ' ' || (i + 1) == val.Length)
                    {
                      if ((i + 1) == val.Length) modelPart.Append(val[i]);
                      if (reg.IsMatch(modelPart.ToString()))
                      {
                        result["speedIndex"] = modelPart.ToString();
                        result["model"] = sb.ToString();
                        break;
                      }
                      else
                      {
                        sb.Append(" ").Append(modelPart.ToString());
                        modelPart.Clear();
                      }
                    }
                    else
                    {
                      modelPart.Append(val[i]);
                    }
                  }
                  if (i == val.Length && !result.ContainsKey("model")) result["model"] = sb.ToString();
                  break;
                case 4:
                  result["spikes"] = sb.ToString().ToUpper() == "Ш";
                  break;
              }
              sb.Clear();
            }
          }

          break;
      }
      return result;
    }

    public decimal getProductPrice(ClosedXML.Excel.IXLRow row)
    {
      var value = row.Cell(parameters["price"].Value<int>()).Value;
      decimal price = 0;
      Decimal.TryParse(value.ToString().Trim(), out price);
      return price;
    }

    public int getProductCount(ClosedXML.Excel.IXLRow row)
    {
      var value = row.Cell(parameters["quantity"].Value<int>()).Value.ToString().Trim();
      int count = value.Contains("Более") ? 20 : 0;
      Int32.TryParse(value, out count);
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

    public string getProductManufacturer(ClosedXML.Excel.IXLRow row)
    {
      throw new NotImplementedException();
    }

    public string getProductModel(ClosedXML.Excel.IXLRow row)
    {
      throw new NotImplementedException();
    }
  }
}
