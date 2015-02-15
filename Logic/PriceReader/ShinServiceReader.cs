using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
    public Dictionary<string, object> readProduct(ClosedXML.Excel.IXLRow row)
    {
      var result = new Dictionary<string, object>();
      result["productId"] = this.getProductId(row);
      result["price"] = this.getProductPrice(row);
      result["count"] = this.getProductCount(row);
      if(String.IsNullOrEmpty(this.getProductId(row))) 
          return result;
      
      var str = row.Cell(4).Value.ToString().Split(' ');

      return result;
    }

    public string getSheetName()
    {
        return parameters["sheet"].ToString();
    }
    public string getProductId(ClosedXML.Excel.IXLRow row)
    {
      return row.Cell(parameters["productId"].Value<int>()).Value.ToString().Trim();
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
      var value = row.Cell(parameters["quantity"].Value<int>()).Value;
      int count = 0;
      Int32.TryParse(value.ToString().Trim(), out count);
      return count;
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
