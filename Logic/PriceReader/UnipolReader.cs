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
  class UnipolReader : IPricelistReader
  {
    private JToken parameters;

    public UnipolReader(JToken parameters)
    {
      this.parameters = parameters;
    }
    public Dictionary<string, object> readProduct(IXLRow row)
    {
      var result = new Dictionary<string, object>();
      if (String.IsNullOrEmpty(this.getProductId(row))) return result;
      result["productId"] = this.getProductId(row);
      result["price"] = this.getProductPrice(row);
      result["count"] = this.getProductCount(row);

      var pType = this.getProductType(row);
      if (pType == ProductType.tyre)
        result["season"] = this.getTyreSeason(row);

      if (String.IsNullOrEmpty(this.getProductId(row)))
        return result;

      var str = this.parseProduct(row.Cell(11), pType);

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
          Regex regex = new Regex(@"[ ]{2,}", RegexOptions.None);
          val = regex.Replace(val, @" ");

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
                  var _obj = sb.ToString().Split('R');
                  var size = _obj[0].Split('/');
                  result["width"] = size[0];
                  result["profile"] = size.Length > 1 ? size[1] : null;
                  result["diameter"] = _obj.Length > 1 ? _obj[1] : null;
                  break;
                case 1:
                  result["speedIndex"] = sb.ToString();
                  break;
                case 2:
                  var modelPart = new StringBuilder();
                  for (i++; i < val.Length; i++) 
                  {
                    if (val[i] == ' ' || (i + 1) == val.Length)
                    {
                      if ((i + 1) == val.Length) modelPart.Append(val[i]);
                      if (modelPart.ToString() == "шип")
                      {
                        result["spikes"] = true;
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
              }
              sb.Clear();
            }
          }
          break;
      }
      return result;
    }

    public decimal getProductPrice(IXLRow row)
    {
      var value = row.Cell(parameters["price"].Value<int>()).Value;
      decimal price = 0;
      Decimal.TryParse(value.ToString().Trim(), out price);
      return price;
    }

    public int getProductCount(IXLRow row)
    {
      var value = row.Cell(parameters["quantity"].Value<int>()).Value.ToString().Trim();
      int count = value.Contains("да") ? 20 : 0;
      Int32.TryParse(value, out count);
      return count;
    }

    public TyreSeason getTyreSeason(IXLRow row)
    {
      var _val = row.Cell(5).Value;
      if (_val != null)
      {
        var _strVal = _val.ToString().ToLower();
        if (_strVal == "лето")
          return TyreSeason.summer;
        if (_strVal == "зима")
          return TyreSeason.winter;
        if (_strVal == "всесезон")
          return TyreSeason.all;
      }
      return TyreSeason.other;
    }

    public string getProductManufacturer(IXLRow row)
    {
      return row.Cell(2).Value.ToString().Trim();
    }

    public string getProductModel(IXLRow row)
    {
      throw new NotImplementedException();
    }
  }
}
