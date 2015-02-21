using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemirPriceBalance.Logic.Product
{
  public enum TyreSeason
  {
    summer,
    winter,
    all,
    other
  }
  class Tyre : Product
  {
    public Tyre()
    {

    }
    public Tyre(string id, string manufacturer, string name, uint quantity, decimal price, string model, TyreSeason season, string profileWidth, string profileHeight, string diameter, string weightIndex, string speedIndex, bool hasSpikes, bool hasRunFlat) 
      : base(id, manufacturer, name, quantity, price)
    {
      this.Model = model;
      this.Season = season;
      this.ProfileWidth = profileWidth;
      this.ProfileHeight = profileHeight;
      this.Diameter = diameter;
      this.WeightIndex = weightIndex;
      this.SpeedIndex = speedIndex;
      this.HasSpikes = hasSpikes;
    }

    public override object[] ToExcelRow()
    {
      return new object[]
      {
        this.Id,
        this.Manufacturer,
        this.Model,
        this.Id,
        this.Season,
        this.ProfileWidth,
        this.ProfileHeight,
        this.Diameter,
        this.WeightIndex,
        this.SpeedIndex,
        this.HasSpikes,
        this.HasRunFlat,
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
        this.Quantity,
        this.Price,
        null,
        null,
        null,
        null,
        null,
        null,
        null
      };
    }
    public string Model { get; set; }
    public TyreSeason Season { get; set; }
    public string ProfileWidth { get; set; }
    public string ProfileHeight { get; set; }
    public string Diameter { get; set; }
    public string WeightIndex { get; set; }
    public string SpeedIndex { get; set; }
    public bool HasSpikes { get; set; }
    public bool HasRunFlat { get; set; }
  }
}
