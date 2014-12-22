using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using System.Threading.Tasks;

namespace DemirPriceBalance.Logic
{
  public enum OfferType
  {
    simple,
    vendorModel
  }

  [Serializable]
  public class Offer
  {
    [XmlElementAttribute(IsNullable = false)]
    public OfferType Type
    {
      get;
      set;
    }
    [XmlElementAttribute(IsNullable = false)]
    public bool Available
    {
      get;
      set;
    }

    public string URL
    {
      get;
      set;
    }
    public string Price
    {
      get;
      set;
    }

    public string CurrencyId
    {
      get;
      set;
    }
    public short CategoryId
    {
      get;
      set;
    }
    public string MarketCategory
    {
      get;
      set;
    }

    public string Picture
    {
      get;
      set;
    }
    public string Vendor
    {
      get;
      set;
    }
    public string Model
    {
      get;
      set;
    }

    public Offer()
    {

    }

    public Offer(OfferType type, bool available, string URL, string price, string currencyId, short categoryId, string marketCategory, string picture, string vendor, string model)
    {
      this.Type = type;
      this.Available = available;
      this.URL = URL;
      this.Price = price;
      this.CurrencyId = currencyId;
      this.CategoryId = categoryId;
      this.MarketCategory = marketCategory;
      this.Picture = picture;
      this.Vendor = vendor;
      this.Model = model;
    }
  }
}
