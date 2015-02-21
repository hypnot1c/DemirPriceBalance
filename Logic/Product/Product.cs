using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemirPriceBalance.Logic.Product
{
  class Product
  {
    public Product()
    {

    }
    public Product(string id, string manufacturer, string name, uint quantity, decimal price)
    {
      this.Id = id;
      this.Manufacturer = Manufacturer;
      this.Name = Name;
      this.Quantity = Quantity;
      this.Price = Price;
    }
    public virtual object[] ToExcelRow()
    {
      return new object[]
      {
        this.Id,
        this.Manufacturer,
        this.Name,
        this.Quantity,
        this.Price
      };
    }
    public string Id { get; set; }
    public string Manufacturer { get; set; }
    public string Name { get; set; }
    public uint Quantity { get; set; }
    public decimal Price { get; set; }
  }
}
