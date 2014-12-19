using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace DemirPriceBalance.Logic
{
  class YML
  {
    public YML(string filePath)
    {
      using (var xml = XmlReader.Create(filePath))
      {
        this.PriceList = XDocument.Load(xml);
      }
    }

    public XDocument PriceList
    {
      get;
      set;
    }


  }
}
