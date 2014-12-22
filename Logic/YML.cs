using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;
using System.IO;

namespace DemirPriceBalance.Logic
{
  class YML
  {
    public YML(string filePath, string outputFile)
    {
      XmlReaderSettings settings = new XmlReaderSettings();
      settings.DtdProcessing = DtdProcessing.Ignore;
      settings.ValidationType = ValidationType.DTD;
      using (var xml = XmlReader.Create(filePath, settings))
      {
        this._outputFile = outputFile;
        this.PriceList = XDocument.Load(xml);
      }
    }

    private string _outputFile;

    public XDocument PriceList
    {
      get;
      set;
    }

    public bool writeOffers(Offer[] offers)
    {
      var root = this.PriceList.Root;
      root.SetAttributeValue("date", DateTime.Now);
      var offersElm = root.Descendants("offers").First();
      var form = new XmlSerializer(typeof(Offer[]));
      var wrt = new StringWriter();
      try
      {
        form.Serialize(wrt, offers);
        offersElm.Add(XElement.Parse(wrt.ToString()));
        offersElm.Save(this._outputFile);
      }
      catch(Exception ex)
      {

      }
      return true;
    }
  }
}
