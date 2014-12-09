using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemirPriceBalance.Logic
{
  class ReadThread
  {
    public EventHandler<ReadFileEventArgs> readDone;
    public string Filename
    {
      get;
      set;
    }

    public string SheetName
    {
      get;
      set;
    }

    public int ColumnID
    {
      get;
      set;
    }

    public int ColumnPrice
    {
      get;
      set;
    }
    public int ColumnCount
    {
      get;
      set;
    }

    public ReadThread(string filename, string sheetName, int clmnId, int clmnPrice, int clmnCount)
    {
      this.Filename = filename;
      this.SheetName = sheetName;
      this.ColumnID = clmnId;
      this.ColumnPrice = clmnPrice;
      this.ColumnCount = clmnCount;
    }

    public void GetDataFromFile(object data)
    {
      var supplierName = (string)data;
      var _data = ExcelReader.readExcel(this.Filename, this.SheetName, this.ColumnID, this.ColumnPrice, this.ColumnCount);
      if (readDone != null)
        readDone(_data, new ReadFileEventArgs(supplierName));
    }

  }
}
