using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using static ExcelAddIn2.MyRibbon;
//using Microsoft.Office.Tools.Excel;


namespace ExcelAddIn2
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //getActiveWorksheet();
            MyRibbon.setCharacteristics();
            MyRibbon.setUrl1("http://127.0.0.1:5000/classify/xcells/api/v0.2/tasks");
            MyRibbon.setUrl2("http://127.0.0.1:5000/classify/xcells/api/v0.2/tasks/run/");

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        public Excel.Worksheet getActiveWorksheet()
        {
            //var WB = Application.Workbooks.Open("Tabelle.xlsx", ReadOnly: false).ActiveSheet;
            return (Excel.Worksheet)Application.ActiveSheet;
            //return WB;
        }
        public Excel.Range getActiveCell() {
            Excel.Range rng = (Excel.Range)Application.ActiveCell;
            return rng;
        }

        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
        
}
