using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemirPriceBalance.Logic
{
  class Config
  {
    public static JObject data;

    public static void LoadConfig()
    {
      Config.data = JObject.Parse(File.ReadAllText("config.json", Encoding.UTF8));
    }
  }
}
