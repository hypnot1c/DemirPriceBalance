using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemirPriceBalance.Logic.Product
{
  class Wheel : Product
  {
    public Wheel()
    {

    }
    public override object[] ToExcelRow(int clmnQnty, int clmnPrice)
    {
      var res = new object[]
      {
        this.Id,
        this.Manufacturer,
        this.Model,
        this.Id,
        this.Diameter,
        this.Width,
        this.DIA,
        this.PCD,
        this.ET,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null
      };
      res[clmnQnty - 1] = this.Quantity;
      res[clmnPrice - 1] = this.Price;
      return res;
    }
    public string Model { get; set; }
    public decimal Diameter { get; set; }
    public decimal Width { get; set; }
    public uint Holes { get; set; }
    public decimal PCD { get; set; }
    public decimal ET { get; set; }
    public decimal DIA { get; set; }
  }
}
