using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemirPriceBalance.Logic.PriceReader
{
  class PriceReaderFactory
  {
    public static IPricelistReader getPriceReader(string provider, JToken parameters)
    {
      switch(provider.ToLower()) 
      {
        case "юнипол":
          return new UnipolReader(parameters);
        case "шинсервис":
          return new ShinServiceReader(parameters);
        case "са.ру":
          return new SaRuReader(parameters);
        default:
          return new ShinServiceReader(parameters);
      }
    }
  }
}
