using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemirPriceBalance.Logic
{
  class PriceList
  {
    public struct Stuff
    {
      public string id;
      public string count;
      public string price;

      public Stuff(string id, string count, string price)
      {
        this.id = id;
        this.count = count;
        this.price = price;
      }
    }

    public PriceList(string supplierName, List<Stuff> goods)
    {
      this.SupplierName = supplierName;
      this.Goods = goods;
    }

    public string SupplierName
    {
      get;
      set;
    }

    public List<Stuff> Goods
    {
      get;
      set;
    }

  }
}
