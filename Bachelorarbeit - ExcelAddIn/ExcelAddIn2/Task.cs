using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn2
{
 

    public class Cell
    {
        public string address { get; set; }
        public string srow { get; set; }
        public string erow { get; set; }
        public string scol { get; set; }
        public string ecol { get; set; }

        public Features features { get; set; } // Features e staroto mycell 
    }

    public class RootObject
    {
        public string id { get; set; }
        public string filename { get; set; }
        public string sheetname { get; set; }
        public List<Cell> cells { get; set; }
    }

    //Root object used only for deserialization 
    public class RootObjectDes
    {
        public List<MyTask> results { get; set; }
    }
    public class CellDes
    {
        public string address { get; set; }
        public string srow { get; set; }
        public string erow { get; set; }
        public string scol { get; set; }
        public string ecol { get; set; }

        public string predicted { get; set; }

        public Features features { get; set; } 
    }

    public class MyTask
    {
        public string taskid { get; set; }
        public string filename { get; set; }
        public string sheetname { get; set; }
        public List<CellDes> cells { get; set; }
    }
    // do not delete
    public class RootObjectId
    {
        public List<RootObject> tasks { get; set; }
    }
}
