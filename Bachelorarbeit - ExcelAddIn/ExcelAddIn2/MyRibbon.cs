using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Net.Http.Headers;
using System.Net;
using System.Net.Http;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.Reflection;
using System.Diagnostics;
using System.Collections;
using System.Threading.Tasks;
using System.Collections.ObjectModel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using static ExcelAddIn2.Form1;
using System.Text.RegularExpressions;

namespace ExcelAddIn2
{
    public partial class MyRibbon
    {
        public static string url1;
        public static string url2;
        public static string baseAddress;
        private static readonly HttpClient client = new HttpClient();
        Stopwatch sw1 = new Stopwatch();
        Stopwatch sw2 = new Stopwatch();
        Stopwatch sw3 = new Stopwatch();

        public static void setUrl1(string url) {
            url1 = url;
        }
        public string getUrl1() {
            return url1;
        }
        public static void setUrl2(string url)
        {
            url2 = url;
        }
        public string getUrl2()
        {
            return url2;
        }
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {


        }  


        private void Send_Click(object sender, RibbonControlEventArgs e)
        {
            Form1 inputForm = new Form1();
            inputForm.ShowDialog();
        }

        private string checkIfCapitalized(Range cell)
        {
            if ( cell != null  && checkType(cell) == "STRING"){
                if (cell.Value != null)
                {
                    Object cellValue = cell.Value;
                    string cellValueString = cellValue.ToString();
                    if (Char.IsUpper(cellValueString[0]))
                        return "1";
                    else
                        return "0";
                }
            }
            return "9";
        }

        private string checkIfStartsWithSpecial(Range cell)
        {
            string specialSymbols = "@#$%^&*/|<>~" + @"\";
            if (cell != null && checkType(cell) == "STRING")
            {
                if (cell.Value != null)
                {
                    Object cellValue = cell.Value;
                    string cellValueString = cellValue.ToString();
                    if (specialSymbols.Contains(cellValueString[0]))
                        return "1";
                    else
                        return "0";
                }
            }
            return "9";
        }
        private string checkIfAlphanumeric(Range cell)
        {

            //  string fullAlphanumericstring = "";
            if (cell != null && checkType(cell) == "STRING")
            {
                if (cell.Value != null)
                {
                    Object cellValue = cell.Value;
                    string cellValueString = cellValue.ToString();
                    
                    if (cellValueString.Any(char.IsDigit) && cellValueString.Any(char.IsLetter) && Regex.IsMatch(cellValueString, @"^[a-zA-Z0-9 ]+$"))
                        return "1";
                    else
                        return "0";
                }
            }
            return "9";
        }
        private string checkIfAlphabetic(Range cell)
        {

            //  string fullAlphanumericstring = "";
            if (cell != null && checkType(cell) == "STRING")
            {
                if (cell.Value != null)
                {
                    Object cellValue = cell.Value;
                    string cellValueString = cellValue.ToString();
                    if (Regex.IsMatch(cellValueString, @"^[a-zA-Z ]+$"))
                        return "1";
                    else
                        return "0";
                }
            }
            return "9";
        }
        private string checkIfNumeric(Range cell)
        {
           // string numericstring = "0123456789 ";
            if (cell != null)
            {
                if (cell.Value != null)
                {
                    Object cellValue = cell.Value;
                    string cellValueString = cellValue.ToString();
                    if (Regex.IsMatch(cellValueString, "^[0-9,.]+$") && cellValueString != "." && cellValueString != ",")
                        return "1";
                    else
                        return "0";
                }
                else
                    return "0";
            }
            return "9";
        }

        private string checkIfFormula(Range cell)
        {
            if (cell != null)
            {
                if(cell.HasFormula)
                    return "1";
                else
                    return "0";
            }
            return "9";
        }
        private string checkIfStartsWithNumber(Range cell)
        {
            string numbers = "0123456789";
            if (cell != null && checkType(cell) == "STRING")
            {
                string cellValue = cell.Value;
                if (numbers.Contains(cellValue[0]))
                    return "1";
                else
                    return "0";
            }
            return "9";
        }
        public class DataObject
        {
            public string Name { get; set; }
        }

        private const string URL = "http://127.0.0.1:5000/classify/xcells/api/v0.2/tasks";
       // private static string urlParameters = "?api_key=123";

        static void Main(string[] args)
        {

        }


      
        public void IterateCells(Worksheet worksheet, Range usedRange)
        {
            
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            List<Cell> cellList = new List<Cell>();



            var objectToSerialize = new RootObject()
            {
                filename = currentSheet.Application.ActiveWorkbook.Name,
                sheetname = currentSheet.Name,
                cells = cellList
            };

            sw1.Start();
            foreach (Range cell in usedRange)
            {
                int r = cell.Row;
                int c = cell.Column;

                if (currentSheet.Cells[r, c].MergeCells)
                {
                    Debug.WriteLine("Ima merge cell sepak -.-");
                }
                else {
                    cellList.Add(new Cell() { address = getCellAddress(cell),
                                            srow = getStartRow(cell).ToString(),
                                            erow = getEndRow(cell).ToString(),
                                            scol = getStartColumn(cell).ToString(),
                                            ecol = getEndColumn(cell).ToString(),          
                                            features = new Features() {
                                                value = makeString(cell.Value2),
                                                FONT_NAME = getFontName(cell),
                                                IS_BOLD = isBold(cell),
                                                IS_ITALIC = isItalic(cell),
                                                IS_STRIKE_OUT = isStrikedOut(cell),
                                                UNDERLINE_TYPE = getUnderlineType(cell),
                                                IS_UNDERLINED = isUnderlined(cell),
                                                OFFSET_TYPE = getOffsetType(cell),
                                                LENGTH = getValueLength(cell),
                                                NUM_OF_TOKENS = getNumberOfTokens(cell),
                                                LEADING_SPACES = getLeadingSpaces(cell),
                                                WIDTH = getCellWidth(cell),
                                                HEIGHT = getCellHeight(cell),
                                                NUMBER_OF_NEIGHBORS = countUsedNeighbours(r, c),
                                                ROW_NUM = getRowOfCell(cell).ToString(),
                                                COLUMN_NUM = getColumnOfCell(cell).ToString(),
                                                TOP_NEIGHBOR_TYPE = getTopNeighborType(r, c),
                                                BOTTOM_NEIGHBOR_TYPE = getBottomNeighborType(r, c),
                                                LEFT_NEIGHBOR_TYPE = getLeftNeighborType(r, c),
                                                RIGHT_NEIGHBOR_TYPE = getRightNeighborType(r, c),
                                                MATCHES_TOP_TYPE = checkIfTypeMatches(cell, getTopNeighborType(r, c)).ToString(),
                                                MATCHES_BOTTOM_TYPE = checkIfTypeMatches(cell, getBottomNeighborType(r, c)).ToString(),
                                                MATCHES_LEFT_TYPE = checkIfTypeMatches(cell, getLeftNeighborType(r, c)).ToString(),
                                                MATCHES_RIGHT_TYPE = checkIfTypeMatches(cell, getRightNeighborType(r, c)).ToString(),
                                                MATCHES_TOP_STYLE = checkIfStyleMatches(cell, getTopNeighbor(r, c)),
                                                MATCHES_BOTTOM_STYLE = checkIfStyleMatches(cell, getBottomNeighbor(r, c)),
                                                MATCHES_LEFT_STYLE = checkIfStyleMatches(cell, getLeftNeighbor(r, c)),
                                                MATCHES_RIGHT_STYLE = checkIfStyleMatches(cell, getRightNeighbor(r, c)),
                                                FONT_COLOR_DEFAULT = isFontColorDefault(getCellFillColor(cell), 0),// 0 is black
                                                FONT_SIZE = getFontSize(cell).ToString(),
                                                IS_CAPITALIZED = checkIfCapitalized(cell),
                                                IS_ALPHANUMERIC = checkIfAlphanumeric(cell),
                                                IS_ALPHABETIC = checkIfAlphabetic(cell),
                                                IS_NUMERIC = checkIfNumeric(cell),
                                                IS_FORMULA = checkIfFormula(cell),
                                                STARTS_WITH_NUMBER = checkIfStartsWithNumber(cell),
                                                STARTS_WITH_SPECIAL = checkIfStartsWithSpecial(cell),
                                                CONTAINS_SPECIAL_CHARS = containsSpecialChars(cell),
                                                CONTAINS_PUNCTUATIONS = containsPunctuations(cell),
                                                WORDS_LIKE_TOTAL = containsWordsLikeTotal(cell),
                                                WORDS_LIKE_TABLE = containsWordsLikeTable(cell),
                                                CONTAINS_COLON = containsColon(cell),
                                                HORIZONTAL_ALIGNMENT = getHorizontalAlignment(cell).ToString(),
                                                VERTICAL_ALIGNMENT = getVerticalAlignment(cell).ToString(),
                                                FILL_PATTERN = getFillPattern(cell).ToString(),
                                                ORIENTATION = getOrientation(cell),
                                                CONTROL = getCellControl(cell),
                                                IS_MERGED = isMerged(cell),
                                                BORDER_TOP_TYPE = getTopBorderType(cell).ToString(),
                                                BORDER_BOTTOM_TYPE = getBottomBorderType(cell).ToString(),
                                                BORDER_LEFT_TYPE = getLeftBorderType(cell).ToString(),
                                                BORDER_RIGHT_TYPE = getRightBorderType(cell).ToString(),
                                                BORDER_TOP_THICKNESS = getTopBorderThickness(cell).ToString(),
                                                BORDER_BOTTOM_THICKNESS = getBottomBorderThickness(cell).ToString(),
                                                BORDER_LEFT_THICKNESS = getLeftBorderThickness(cell).ToString(),
                                                BORDER_RIGHT_THICKNESS = getRightBorderThickness(cell).ToString(),
                                                INDENTATIONS = getIndentations(cell).ToString(),
                                                NUM_OF_BORDERS = countBorders(cell).ToString(),
                                                NUMBER_OF_CELLS = getNumberOfCells(cell).ToString(),
                                                FORMULA_VAL_TYPE = getReturnTypeOfFormula(cell),
                                                IS_AGGREGATION_FORMULA = checkIfCellContainsAggregationFormula(cell),
                                                REF_VAL_TYPE = getRefValValue(cell),
                                                REF_IS_AGGREGATION_FORMULA = checkIfRefIsAggregationFormula(cell),
                                                IN_YEAR_RANGE = inYearRange(cell),
                                                IS_UPPER_CASE = isUppercase(cell),
                                                FILL_COLOR_DEFAULT = isFontColorDefault(getCellFillColor(cell), 0),
                                                //FONT_COLOR = getFontColor(cell)


                                            }
                    });
                    Debug.WriteLine(" ************************** CELL " + cell.Address + " DOWN ");
                }
                
            }
  

 
            List <Range> lstWithFirstElements = new List<Range>();
            foreach (Range cell in usedRange)
            {
                int r = cell.Row;
                int c = cell.Column;
                int i = 0;
                List<Range> lst = new List<Range>();
     
                if (currentSheet.Cells[r, c].MergeCells)
                {
                    foreach (Range ra in currentSheet.Cells[r, c].MergeArea) {
                        lst.Add(ra);
                    }

                    lstWithFirstElements.Add(lst.ElementAt(0));

                 
                }
               

            }
            Debug.WriteLine("Golemata lista imase volku elementi: " + lstWithFirstElements.Count);

            // eliminate the duplicates from lstWithFirstElements
            for(int i = 0; i < lstWithFirstElements.Count; i ++)
            {

                // i + 1 <= lstWithFirstElements.Count 
                if (lstWithFirstElements.ElementAtOrDefault(i+1) != null)
                {
                    if (Globals.ThisAddIn.Application.Intersect(usedRange, lstWithFirstElements.ElementAt(i + 1).MergeArea) != null)
                    { 
                        Debug.WriteLine("Adresa na  element:  " + lstWithFirstElements.ElementAt(i).MergeArea.Address);
                    Debug.WriteLine("Adresa na next element:" + lstWithFirstElements.ElementAt(i + 1).MergeArea.Address);

                    if (lstWithFirstElements.ElementAt(i).MergeArea.Address == lstWithFirstElements.ElementAt(i + 1).MergeArea.Address)
                        lstWithFirstElements.Remove(lstWithFirstElements.ElementAt(i + 1));

                    if ((lstWithFirstElements.Count - 1 > 0) && (lstWithFirstElements.Count - 2) > 0)
                        if (lstWithFirstElements.ElementAt(lstWithFirstElements.Count - 1).Address == lstWithFirstElements.ElementAt(lstWithFirstElements.Count - 2).Address)
                            lstWithFirstElements.Remove(lstWithFirstElements.ElementAt(lstWithFirstElements.Count - 1));
                    }
                }  
            }
              

            foreach (Range cell in lstWithFirstElements) {

                int r = cell.Row;
                int c = cell.Column;

                cellList.Add(new Cell()
                {
                    address = getCellAddress(cell),
                    srow = getStartRow(cell).ToString(),
                    erow = getEndRow(cell).ToString(),
                    scol = getStartColumn(cell).ToString(),
                    ecol = getEndColumn(cell).ToString(),
                    features = new Features()
                    {
                      //  address = getCellAddress(cell),
                        value = makeString(cell.Value2),
                        //value = getCellValue(cell),
                      //  ROW_NUM = getRowOfCell(cell).ToString(),
                      //  COLUMN_NUM = getColumnOfCell(cell).ToString(),
                        NUMBER_OF_NEIGHBORS = countUsedNeighbours(r, c),
                        WIDTH = getCellWidth(cell),
                        HEIGHT = getCellHeight(cell),
                        TOP_NEIGHBOR_TYPE = getTopNeighborType(r, c),
                        BOTTOM_NEIGHBOR_TYPE = getBottomNeighborType(r, c),
                        LEFT_NEIGHBOR_TYPE = getLeftNeighborType(r, c),
                        RIGHT_NEIGHBOR_TYPE = getRightNeighborType(r, c),
                        MATCHES_TOP_TYPE = checkIfTypeMatches(cell, getTopNeighborType(r, c)).ToString(),
                        MATCHES_BOTTOM_TYPE = checkIfTypeMatches(cell, getBottomNeighborType(r, c)).ToString(),
                        MATCHES_LEFT_TYPE = checkIfTypeMatches(cell, getLeftNeighborType(r, c)).ToString(),
                        MATCHES_RIGHT_TYPE = checkIfTypeMatches(cell, getRightNeighborType(r, c)).ToString(),
                        MATCHES_TOP_STYLE = checkIfStyleMatches(cell, getTopNeighbor(r, c)),
                        MATCHES_BOTTOM_STYLE = checkIfStyleMatches(cell, getBottomNeighbor(r, c)),
                        MATCHES_LEFT_STYLE = checkIfStyleMatches(cell, getLeftNeighbor(r, c)),
                        MATCHES_RIGHT_STYLE = checkIfStyleMatches(cell, getRightNeighbor(r, c)),
                        FONT_COLOR_DEFAULT = isFontColorDefault(getCellFillColor(cell), 0),// 0 is black
                        FONT_SIZE = getFontSize(cell).ToString(),
                        IS_BOLD = isBold(cell),
                        IS_ITALIC = isItalic(cell),
                        IS_STRIKE_OUT = isStrikedOut(cell),
                        UNDERLINE_TYPE = getUnderlineType(cell),
                        IS_UNDERLINED = isUnderlined(cell),
                        OFFSET_TYPE = getOffsetType(cell),
                        LENGTH = getValueLength(cell),
                        NUM_OF_TOKENS = getNumberOfTokens(cell),
                        LEADING_SPACES = getLeadingSpaces(cell),
                        IS_CAPITALIZED = checkIfCapitalized(cell),
                        CONTAINS_SPECIAL_CHARS = containsSpecialChars(cell),
                        CONTAINS_PUNCTUATIONS = containsPunctuations(cell),
                        WORDS_LIKE_TOTAL = containsWordsLikeTotal(cell),
                        WORDS_LIKE_TABLE = containsWordsLikeTable(cell),
                        HORIZONTAL_ALIGNMENT = getHorizontalAlignment(cell).ToString(),
                        VERTICAL_ALIGNMENT = getVerticalAlignment(cell).ToString(),
                        FILL_PATTERN = getFillPattern(cell).ToString(),
                        ORIENTATION = getOrientation(cell),
                        CONTROL = getCellControl(cell),
                        IS_MERGED = isMerged(cell),
                        BORDER_TOP_TYPE = getTopBorderType(cell).ToString(),
                        BORDER_BOTTOM_TYPE = getBottomBorderType(cell).ToString(),
                        BORDER_LEFT_TYPE = getLeftBorderType(cell).ToString(),
                        BORDER_RIGHT_TYPE = getRightBorderType(cell).ToString(),
                        BORDER_TOP_THICKNESS = getTopBorderThickness(cell).ToString(),
                        BORDER_BOTTOM_THICKNESS = getBottomBorderThickness(cell).ToString(),
                        BORDER_LEFT_THICKNESS = getLeftBorderThickness(cell).ToString(),
                        BORDER_RIGHT_THICKNESS = getRightBorderThickness(cell).ToString(),
                        INDENTATIONS = getIndentations(cell).ToString(),
                        NUM_OF_BORDERS = countBorders(cell).ToString(),
                        NUMBER_OF_CELLS = getNumberOfCells(cell).ToString(),
                        FORMULA_VAL_TYPE = getReturnTypeOfFormula(cell),
                        IS_AGGREGATION_FORMULA = checkIfCellContainsAggregationFormula(cell),
                        REF_VAL_TYPE = getRefValValue(cell),
                        REF_IS_AGGREGATION_FORMULA = checkIfRefIsAggregationFormula(cell),
                        IN_YEAR_RANGE = inYearRange(cell),
                        IS_UPPER_CASE = isUppercase(cell),
                        FILL_COLOR_DEFAULT = isFontColorDefault(getCellFillColor(cell), 0),
                        //FONT_COLOR = getFontColor(cell)


                    }
                });
                Debug.WriteLine(" ************************** CELL " + cell.Address + " DOWN ");
            }
            //Debug.WriteLine("Number of merged cells: " + lstWithFirstElements.Count);
            File.AppendAllText(@"c:\Users\ljti\Desktop\newFile.json", JsonConvert.SerializeObject(objectToSerialize, Formatting.Indented));

        }

        public string inYearRange(Range cell)
        {
            if (cell != null) {
                object cellValue = cell.Value;
                if (cellValue != null) { 
                Type t = cellValue.GetType();
                if (t.Equals(typeof(double)))
                {
                    double doubleValue = cell.Value;
                    if (doubleValue % 1 >= 0)
                    {
                        if (doubleValue >= 1900 && doubleValue <= 2000)
                            return "1";
                        else
                            return "0";
                    }
                    else return "9";   
                }
                }
                return "9";
            }
            return "9";
        }
        public string isUppercase(Range cell)
        {
            if (cell != null)
            {
                
                object cellValue = cell.Value;
                if (cellValue != null) { 
                Type t = cellValue.GetType();
                if (t.Equals(typeof(string)))
                {
                    string myString = cell.Value;
                    if (myString.All(char.IsUpper))
                        return "1";
                    else
                        return "0";

                }
                }
                return "9";
            }
            return "9";
        }


        public int getStartRow(Range cell) {
            if (cell != null){
               return cell.Row;
            }
            return 9;
        }

        public int getEndRow(Range cell)
        {
            if (cell != null)
            {
                if (cell.MergeCells) {
                    int endRow = cell.Row + cell.Rows.Count -1 ;
                    return endRow;
                }
                 
                else return cell.Row;
            }
            return 9;
        }
        public int getStartColumn(Range cell)
        {
            if (cell != null)
            {
                return cell.Column;
            }
            return 9;
        }
        public int getEndColumn(Range cell)
        {
            if (cell != null)
            {
                if (cell.MergeCells)
                {
                    int endColumn = cell.Column + cell.Columns.Count -1;
                    return endColumn;
                }
                else return cell.Column;
            }
            return 9;
        }

        //classify

        private void sendToAPI_Click(object sender, RibbonControlEventArgs e)
        {

            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            IterateCells(currentSheet, currentSheet.UsedRange);

            RunAsync(getUrl1(), getUrl2()).Wait();

        }

        public static void setCharacteristics() {
            //client.BaseAddress = new Uri(path);
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        }

        public static async Task RunAsync(string pathPost, string pathGet)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
       
           // client.BaseAddress = new Uri("http://127.0.0.1:5000/classify/xcells/api/v0.2/tasks");
           // client.DefaultRequestHeaders.Accept.Clear();
           // client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            try
            {
                string jsonString = loadJson();
                RootObject rootObj = JsonConvert.DeserializeObject<RootObject>(jsonString);

                var url = await CreateProductAsync(pathPost, rootObj);

                //empty json file
                System.IO.File.WriteAllText(@"c:\Users\ljti\Desktop\newFile.json", string.Empty);

                //get the current id of task
                int nextId = await getNextId(pathPost);
             
                RootObject receivedRootObject = await GetTaskAsync(pathGet + nextId);

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            Console.ReadLine();
        }


        public static async Task<int> getNextId(string path)
        {

            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            int nextId = 0;

            HttpResponseMessage response = await client.GetAsync(path);
            var responseBody = response.Content.ReadAsStringAsync().Result;

            RootObjectId allTasks = JsonConvert.DeserializeObject<RootObjectId>(responseBody);

            foreach (var item in allTasks.tasks)
            {
                nextId = Convert.ToInt32(item.id);
            }
            return nextId;
        }

        static public string loadJson()
        {
        
            string json = "";
            using (StreamReader reader = new StreamReader(@"c:\Users\ljti\Desktop\newFile.json"))
            {
                json = reader.ReadToEnd();
                dynamic files = JsonConvert.DeserializeObject(json);
            }
            return json;
        }

      

        public static async Task<RootObject> GetTaskAsync(string path)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            RootObject product = null;
            HttpResponseMessage response = await client.GetAsync(path);

            var responseBody = response.Content.ReadAsStringAsync().Result;

            RootObjectDes data = JsonConvert.DeserializeObject<RootObjectDes>(responseBody);

            if (response.IsSuccessStatusCode)
            {
                product = await response.Content.ReadAsAsync<RootObject>();
            }

            foreach (var item in data.results) {

                foreach (var cell in item.cells) {

                    Range rng = currentSheet.Range[cell.address.ToString()];

                    System.Drawing.Color color1 = System.Drawing.Color.FromArgb(204, 255, 204); // green
                    System.Drawing.Color color2 = System.Drawing.Color.FromArgb(180, 205, 205); //silver
                    System.Drawing.Color color3 = System.Drawing.Color.FromArgb(254, System.Drawing.Color.LightCoral); //coral
                    System.Drawing.Color color4 = System.Drawing.Color.FromArgb(255, 255, 179); //yellow
                    System.Drawing.Color color5 = System.Drawing.Color.FromArgb(213, 186, 219); // violet

                    if (cell.predicted == "data")
                       rng.Interior.Color = System.Drawing.ColorTranslator.ToOle(color1);

                    if (cell.predicted == "header")
                        rng.Interior.Color = System.Drawing.ColorTranslator.ToOle(color2);

                    if (cell.predicted == "derived")
                        rng.Interior.Color = System.Drawing.ColorTranslator.ToOle(color3);

                    if (cell.predicted == "metadata")
                        rng.Interior.Color = System.Drawing.ColorTranslator.ToOle(color4);

                    if (cell.predicted == "attributes")
                        rng.Interior.Color = System.Drawing.ColorTranslator.ToOle(color5);
                }
            }
            return product;
        }

      


      public static async Task<Uri> CreateProductAsync(string url, RootObject product)
        {
           
            HttpResponseMessage response = await client.PostAsJsonAsync(url, product);
            Console.WriteLine("camu sme tamu sme  ");
            // response.EnsureSuccessStatusCode();

            return response.Headers.Location;
        }
     
        

        // with consideration od merged cells 
        public string getCellValue(Range cell) {

            string cellValue = "";
            object obj;

            if (cell != null)
            {
                obj = cell.Value;
                cellValue = obj.ToString();

            
                if (cell.MergeCells) {
                    obj = cell.MergeArea.Cells[1, 1].Value;
                    cellValue = obj.ToString();
                }

                return cellValue;
     
            }
            return "9";

        }

        public string getCellAddress(Range cell) {

            if (cell != null) {
                if (cell.MergeCells)
                    return cell.MergeArea.Address;
                return cell.Address;
            }
            return "9";
        }
        public string countUsedNeighbours(int r, int c)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            Range usedRange = currentSheet.UsedRange;
            int counter = 0;


            if (r - 1 > 0)
            {
                if (currentSheet.Cells[r - 1, c].Value2 != null)
                    counter++;
            }
            if (c - 1 > 0)
            {
                if (currentSheet.Cells[r, c - 1].Value2 != null)
                    counter++;
            }


            if (currentSheet.Cells[r + 1, c].Value2 != null)
                counter++;

            if (currentSheet.Cells[r, c + 1].Value2 != null)
                counter++;

            return counter.ToString();
        }

        public string getLabel(String address)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();

            if (address != "") {

                string newString = address.Replace("$","");
                return newString;
            }
            else
                return ""; 
        }
        public string makeString(Object value)
        {
            if (value != null)
                return value.ToString();
            else
                return "";
        }

        enum contentType { NUMERIC, STRING, BOOLEAN, DATE, FORMULA };
        public string checkType(Range rng)
        {
            if (rng.Value == null)
            {
                return "9";
            }
            Object value = rng.Value;
            Type t = value.GetType();

            if (rng.HasFormula)
                return contentType.FORMULA.ToString();
            if (t.Equals(typeof(int)))
                return contentType.NUMERIC.ToString();
            if (t.Equals(typeof(double)))
                return contentType.NUMERIC.ToString();
            if (t.Equals(typeof(float)))
                return contentType.NUMERIC.ToString();
            if (t.Equals(typeof(String)))
                return contentType.STRING.ToString();
            if (t.Equals(typeof(Boolean)))
                return contentType.BOOLEAN.ToString();
            if (t.Equals(typeof(DateTime)))
                return contentType.DATE.ToString();

            return contentType.STRING.ToString();


        }
        public string getTypeValueInQuotes(Range cell) {
            if (cell.Value != null)
            {
                Object cellValue = cell.Value;
                Type t = cellValue.GetType();
                if (t.Equals(typeof(String))) {
                    string cellV = cell.Value;
                    int stringSize = cellV.Count();
                if ((cellV[0] == '"' && cellV[stringSize-1] == '"') || (cellV[0] == '\'' && cellV[stringSize-1] == '\''))
                {
                    cellV = cellV.Trim('"');
                        if (Regex.IsMatch(cellV, @"^[1-9]\d{0,2}(\.\d{3})*(,\d+)%?$"))
                            return "NUMERIC";
                        else
                            return checkType(cell);
                }
                }
                else
                    return checkType(cell);

            }
            return "9";
            
        }
        enum formulaValueType { NUMERIC, STRING, BOOLEAN, DATE};

        public string checkFormulaType(Range rng)
        {
            if (rng.Value == null)
            {
                return "9";
            }
            Object value = rng.Value;
            Type t = value.GetType();

            if (t.Equals(typeof(int)))
                return formulaValueType.NUMERIC.ToString();
            if (t.Equals(typeof(double)))
                return formulaValueType.NUMERIC.ToString();
            if (t.Equals(typeof(float)))
                return formulaValueType.NUMERIC.ToString();
            if (t.Equals(typeof(String)))
                return formulaValueType.STRING.ToString();
            if (t.Equals(typeof(Boolean)))
                return formulaValueType.BOOLEAN.ToString();
            if (t.Equals(typeof(DateTime)))
                return formulaValueType.DATE.ToString();



            return formulaValueType.STRING.ToString();


        }


        public string getTopNeighborType(int r, int c)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();

            if (currentSheet.Cells[r, c].MergeCells)
            {
                getTopNeighborTypeOfMergedCell(r, c);
            }
            if (r - 1 > 0)
            {
                if (currentSheet.Cells[r - 1, c].Value2 != null)
                    return checkType(currentSheet.Cells[r - 1, c]);
            }
            else
                return "9"; 
            return "9";

        }
        public string getBottomNeighborType(int r, int c)
        {

            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            if (currentSheet.Cells[r, c].MergeCells)
            {
                getBottomNeighborTypeOfMergedCell(r, c);
            }
            if (currentSheet.Cells[r + 1, c].Value2 != null)
                return checkType(currentSheet.Cells[r + 1, c]);
            else
                return "9";
        }
        public string getLeftNeighborType(int r, int c)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();

            if (currentSheet.Cells[r, c].MergeCells)
            {
                getLeftNeighborTypeOfMergedCell(r, c);
            }
            if (c - 1 > 0)
            {
                if (currentSheet.Cells[r, c - 1].Value2 != null)
                    return checkType(currentSheet.Cells[r, c - 1]);
            }
            else
                return "9";
            return "9";
        }
        public string getRightNeighborType(int r, int c)
        {

            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();

            if (currentSheet.Cells[r, c].MergeCells)
            {
                getRightNeighborTypeOfMergedCell(r, c);
            }
            if (currentSheet.Cells[r, c + 1].Value2 != null)
                return checkType(currentSheet.Cells[r, c + 1]);
            else
                return "9";
        }


        public string getTopNeighborTypeOfMergedCell(int r, int c)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            int first = 0;
            int  last = 0;

            Range rng;
     
            if (currentSheet.Cells[r, c].MergeCells)
            {
               
                rng = currentSheet.Cells[r, c].MergeArea;

               
                if (r - 1 > 0)
                {
                    first = rng.Column;
                    last = rng.Column + rng.Columns.Count;

                    for (int i = first; i <= last; i++) {
                       
                        if (currentSheet.Cells[r - 1, i].Value != null) 
                        {
                            string s = currentSheet.Cells[r - 1, i].Address;
                            return checkType(currentSheet.Cells[r - 1, i]);
                        }
                           
                    }
     
    
                    }
                        return "9";
                    
                    
                }

                return "9";
          
        }

       
        public string getBottomNeighborTypeOfMergedCell(int r, int c){
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            int first = 0;
            int last = 0;

            Range rng;
          
            if (currentSheet.Cells[r, c].MergeCells)
            {
      
                rng = currentSheet.Cells[r, c].MergeArea;

                string str = currentSheet.Cells[r + 1, c].Address;
              

               
                if (currentSheet.Cells[r+1, c] != null)
                {
                   
                    first = rng.Column;
                    last = rng.Columns.Count + rng.Column - 1 ;

                   
                    int newRow = rng.Rows.Count + rng.Row ;

                   
                    for (int i = first; i <= last; i++)
                    {
                        
                        if (currentSheet.Cells[newRow, i].Value != null) 
                        {
                            string s = currentSheet.Cells[newRow, i].Address;
                          
                            return checkType(currentSheet.Cells[newRow, i]);

                        }
                     
                    }
                }
                return "9";
             }
            return "9";

           
        }

        public string getLeftNeighborTypeOfMergedCell(int r, int c)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            int first = 0;
            int last = 0;

            Range rng;

        
            if (currentSheet.Cells[r, c].MergeCells)
            {
               
                rng = currentSheet.Cells[r, c].MergeArea;

             

               if (c - 1 > 0)
                {
                    first = rng.Row;
                    last = rng.Rows.Count + rng.Row;

         
                

                    for (int i = first; i <= last; i++)
                    {
                  
                        if (currentSheet.Cells[i , c-1].Value != null) 
                        {
                            string s = currentSheet.Cells[i, c-1].Address;

                            // return checkType(currentSheet.Cells[r + 1, i]);
                            return checkType(currentSheet.Cells[i , c - 1]);

                        }

                    }
                }
                return "9";
            }
            return "9";


        }
        public string getRightNeighborTypeOfMergedCell(int r, int c)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            int first = 0;
            int last = 0;

            Range rng;

         
            if (currentSheet.Cells[r, c].MergeCells)
            {
              
                rng = currentSheet.Cells[r, c].MergeArea;   

                if (currentSheet.Cells[r, c + 1].Value2 != null)
                {
                    first = rng.Row;
                    last = rng.Rows.Count + rng.Row;


                   

                    for (int i = first; i <= last; i++)
                    {
                  
                        if (currentSheet.Cells[i, c + 1].Value != null) 
                        {
                            string s = currentSheet.Cells[i, c + 1].Address;
                        
                            return checkType(currentSheet.Cells[i, c + 1]);

                        }
                      
                    }
                }
                else
                    return "9";
            }
                return "9";
            }

        // for merged cells; based on row and column number check if the merged cell is already contained in the .json file
        public string checkIfCellAlreadyExists(Range cell)
        {

            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            Range usedRange = currentSheet.UsedRange;

            var dict = new Dictionary<string, int>();

       
            foreach (Range r in usedRange)
            {

                if (r.MergeCells)
                {
                    Debug.WriteLine("flufli1 ");
                    if (dict.ContainsKey(r.MergeArea.Address))
                    {
                        Debug.WriteLine("vo if sme i adresa e " + r.Address);
                        dict[r.MergeArea.Address]++;
                    }
                    else
                    {
                        Debug.WriteLine("vo else sme i se dodade " + r.MergeArea.Address);
                        dict.Add(r.MergeArea.Address, 1);
                    }

                    Debug.WriteLine(" dict[r.MergeArea.Address]1 e  " + dict[r.MergeArea.Address]);
                    if (dict.ContainsKey(r.Address))
                    {
                        if (dict[cell.MergeArea.Address] > 1)
                            return "1";
                        else
                            return "0";
                    }
                }
                else
                {
                    Debug.WriteLine("flufli2 ");
                    if (dict.ContainsKey(r.Address))
                    {
                        Debug.WriteLine("vo if sme i adresa e " + r.Address);
                        dict[r.Address]++;
                    }
                    else
                    {
                        Debug.WriteLine("vo else sme i se dodade " + r.Address);
                        dict.Add(r.Address, 1);
                    }

                    Debug.WriteLine(" dict[r.Address]2 e  " + dict[r.Address]);

                    // ako postoi vo dict
                    if(dict.ContainsKey(cell.Address)) {
                        if (dict[cell.Address] > 1)
                            return "1";
                        else
                            return "0";

                    }

                }

            }
            return "9";
        }

        public int checkIfTypeMatches(Range cell1, string typeOfSecondCell)
        {

            if (typeOfSecondCell == "9")
                return 9;
            if (checkType(cell1) == typeOfSecondCell)
                return 1;
            else
                return 0;
        }
        public double getCellFillColor(Range cell)
        {


            if (cell != null)
                return cell.Interior.Color;
            else
                return 0;

        }
        public double getCellFontColor(Range cell)
        {
            if (cell != null) {
                if (cell.Characters.Font.Color.Equals(typeof(DBNull)) == false)
                    return cell.Characters.Font.Color;
                else
                    return 0;
            }
            else
                return 0;

        }

        // FONT_COLOR_DEFAULT
        public string isFontColorDefault(double cellColor, double defaultColor)
        {

            if (cellColor == defaultColor)
                return "1";
            else
                return "0";

        }
        // FILL_COLOR_DEFAULT
        public string isFillColorDefault(double cellColor, double defaultColor)
        {

            if (cellColor == defaultColor)
                return "1";
            else
                return "0";

        }
        public double getFontSize(Range cell)
        {
            if (cell != null)
                return cell.Characters.Font.Size;
            else
                return 0;
        }
        public string getFontColor(Range cell) {
            if (cell != null)
            {
                double fontColor = cell.Characters.Font.Color;
                return fontColor.ToString();
            }
            
            else
                return "0";
        } 
        public string isBold(Range cell)
        {
            if (cell != null)
            {
                if (cell.Characters.Font.Bold)
                    return "1";
                else
                    return "0";
            }
            return "9";
        }
        public string isItalic(Range cell)
        {
            if (cell != null)
            {
                if (cell.Characters.Font.Italic)
                    return "1";
                else
                    return "0";
            }
            return "9";
        }
        public string isStrikedOut(Range cell)
        {
            if (cell != null)
            {
                if (cell.Characters.Font.Strikethrough)
                    return "1";
                else
                    return "0";
            }
            return "9";
        }
        public string getUnderlineType(Range cell)
        {
            if (cell != null)
            {

                int i = cell.Font.Underline;
                if (i == 2) // XlUnderlineStyle.xlUnderlineStyleSingle.  ActiveCell.Font.Underline = xlUnderlineStyleSingle
                    return "single";
                if (i == -4119)
                    return "double";
                if (i == 4)
                    return "single accounting";
                if (i == 5)
                    return "double accounting";
                if (i == -4142)
                    return "none";
            }

            return "9";

        }

        public string isUnderlined(Range cell)
        {
            if (cell != null)
            {

                int i = cell.Font.Underline;
                if (i == -4142)
                    return "no";
                else
                    return "yes";
            }

            return "";

        }
        public string getOffsetType(Range cell)
        {
            if (cell != null)
            {
                if (cell.Characters.Font.Superscript)
                    return "superscript";
                if (cell.Characters.Font.Subscript)
                    return "subscript";
                if (cell.Characters.Font.Superscript == false && cell.Characters.Font.Subscript == false)
                    return "none";
            }
            return "";

        }


        //style features 

        public string getOrientation(Range cell)
        {

            if (cell != null)
            {

                if (cell.Orientation == 0)
                    return "0";
                else
                    return "1";
            }
            else return "9";
        }


        public string getCellControl(Range cell)
        {
            if (cell != null)
            {
                if (cell.WrapText && cell.ShrinkToFit)
                    return "1";
                else
                    return "0";
            }
            return "9";
        }

        public int assignValueToLineStyle(int lineStyle) {

            if (lineStyle == -4142) //LineStyleNone
            {
                return 0;
            }
                
            if (lineStyle == 1) // Continuous
            {
               
                return 1;
            }
            if (lineStyle == -4118) // Dot
            {
                
                return 2;
            }
            
            if (lineStyle == -4115) // Dash
            {
              
                return 3;
            }
            if (lineStyle == 4) // DashDot
            {
               
                return 4;
            }
        
            if (lineStyle == 5) // DashDotDot
            {
               
                return 5;
            }
          
            if (lineStyle == -4119) // Double 
            {
                
                return 6;
            }
            if (lineStyle == 13)  // SlantDashDot
            {
               
                return 7;
            }

            return 9;
        }
        public int assignThicknessToBorder(int weight)
        {
            if (weight == 2) // thin = normal 
                return 0;
            if (weight == 1) //hairline
                return 1;
            if (weight == -4138) //medium
                return 2;
            if (weight == 4) //thick
                return 3;
            return 0;
        }
        public int getTopBorderThickness(Range cell)
        {
            if (cell != null)
            {
                if (cell.Borders[XlBordersIndex.xlEdgeTop] != null)
                {
                    return assignThicknessToBorder(cell.Borders[XlBordersIndex.xlEdgeTop].Weight);
                    //return cell.Borders[XlBordersIndex.xlEdgeTop].LineStyle;
                }
                else return 9;

            }
            return 9;

        }
        public int getBottomBorderThickness(Range cell)
        {
            if (cell != null)
            {
                if (cell.Borders[XlBordersIndex.xlEdgeBottom] != null)
                {
                    return assignThicknessToBorder(cell.Borders[XlBordersIndex.xlEdgeBottom].Weight);
                    //return cell.Borders[XlBordersIndex.xlEdgeTop].LineStyle;
                }
                else return 9;

            }
            return 9;

        }
        public int getLeftBorderThickness(Range cell)
        {
            if (cell != null)
            {
                if (cell.Borders[XlBordersIndex.xlEdgeLeft] != null)
                {
                    return assignThicknessToBorder(cell.Borders[XlBordersIndex.xlEdgeLeft].Weight);
                    //return cell.Borders[XlBordersIndex.xlEdgeTop].LineStyle;
                }
                else return 9;

            }
            return 9;

        }
        public int getRightBorderThickness(Range cell)
        {
            if (cell != null)
            {
                if (cell.Borders[XlBordersIndex.xlEdgeRight] != null)
                {
                    return assignThicknessToBorder(cell.Borders[XlBordersIndex.xlEdgeRight].Weight);
                    //return cell.Borders[XlBordersIndex.xlEdgeTop].LineStyle;
                }
                else return 9;

            }
            return 9;

        }

        public int getIndentations(Range cell) {
            if (cell != null)
                return cell.IndentLevel;
            else return 9;
        }

        public int getTopBorderType(Range cell)
        {
            if (cell != null)
            {
                if (cell.Borders[XlBordersIndex.xlEdgeTop] != null)
                {
                    return assignValueToLineStyle(cell.Borders[XlBordersIndex.xlEdgeTop].LineStyle);   
                    //return cell.Borders[XlBordersIndex.xlEdgeTop].LineStyle;
                }
                else return 9;

            }
            return 9;

        }
        public int getBottomBorderType(Range cell)
        {
            if (cell != null)
            {
                if (cell.Borders[XlBordersIndex.xlEdgeBottom] != null)
                {
                    return assignValueToLineStyle(cell.Borders[XlBordersIndex.xlEdgeBottom].LineStyle);
                   // return cell.Borders[XlBordersIndex.xlEdgeBottom].LineStyle;
                }
                 
                else return 9;

            }
            return 9;

        }
        public int getLeftBorderType(Range cell)
        {
            if (cell != null)
            {
                if (cell.Borders[XlBordersIndex.xlEdgeLeft] != null)
                {
                    return assignValueToLineStyle(cell.Borders[XlBordersIndex.xlEdgeLeft].LineStyle);
                    //return cell.Borders[XlBordersIndex.xlEdgeLeft].LineStyle;
                }

                else return 9;

            }
            return 9;

        }
        public int getRightBorderType(Range cell)
        {
            if (cell != null)
            {
                if (cell.Borders[XlBordersIndex.xlEdgeRight] != null)
                {
                    return assignValueToLineStyle(cell.Borders[XlBordersIndex.xlEdgeRight].LineStyle);
                    //return cell.Borders[XlBordersIndex.xlEdgeRight].LineStyle;
                }

                else return 9;

            }
            return 9;

        }

        //NUM_OF_BORDERS
        public int countBorders(Range cell) {
            int i = 0;
            if (cell != null) {
                if (assignValueToLineStyle(cell.Borders[XlBordersIndex.xlEdgeTop].LineStyle) != 0)
                    i++;
                if (assignValueToLineStyle(cell.Borders[XlBordersIndex.xlEdgeBottom].LineStyle) != 0)
                    i++;
                if (assignValueToLineStyle(cell.Borders[XlBordersIndex.xlEdgeLeft].LineStyle) != 0)
                    i++;
                if (assignValueToLineStyle(cell.Borders[XlBordersIndex.xlEdgeRight].LineStyle) != 0)
                    i++;
                return i;
            }
           
            return 0;

        }
        public string getFontName(Range cell) {

            if (cell != null)
            {
                if (cell.Value != null)
                {
                    return cell.Font.Name;
                }
            }
            return "9";
        }
        public string getValueLength(Range cell)
        {
            if (cell != null)
            {
                if (cell.Value != null)
                {
                    Object i = cell.Value;
                    string myString = i.ToString();
                    int j = myString.ToList<char>().Count;
                    return j.ToString();
                }
            }
            return "9";

        }

        public string getNumberOfTokens(Range cell)
        {
            if (cell != null)
            {
                if (cell.Value != null)
                {
                    Object cellValue = cell.Value;
                    string myString = cellValue.ToString();
                    int counter = 1;

                    if (myString.Contains(""))
                    {
                        foreach (char c in myString.ToList())
                        {
                            if (c == ' ')
                                counter++;
                        }
                    }
                    return counter.ToString();
                }
            }
            return "9";

        }

        public string getLeadingSpaces(Range cell)
        {
            if (cell != null)
            {
                if (cell.Value != null)
                {
                    Object i = cell.Value;
                    string myString = i.ToString();
                    int counter = 0;
                    int m = 1;

                   
                    if (myString.First<char>() == ' ')
                    {
                        if (myString.Length > 1)
                        {

                            while (myString.ElementAt<char>(m) == ' ')
                        {
                            m++;
                            counter++;
                        }
                            counter++;
                     }
                    }




                    return counter.ToString();
                }
            }
            return "";

        }

        public string containsSpecialChars(Range cell)
        {
            if (cell != null)
            {
                if (cell.Value != null)
                {
                     Object i = cell.Value;
                     string myString = i.ToString();
                    // if (myString.Any(ch => !Char.IsLetterOrDigit(ch)))
                    if (myString.Contains("@") || myString.Contains("#") || myString.Contains("$") || myString.Contains("%") || myString.Contains("^") || myString.Contains("&") || myString.Contains("*")
                        || myString.Contains("/") || myString.Contains("|") || myString.Contains("<") || myString.Contains(">") || myString.Contains("~") || myString.Contains(@"\"))
                        return "1";
                    else
                        return "0";
                }
                //@#$%^&*/|<>~\

            }
            return "";

        }

        public string containsPunctuations(Range cell)
        {
            if (cell != null)
            {
                if (cell.Value != null)
                {
                    Object i = cell.Value;
                    string myString = i.ToString();
                    if (myString.Contains(";") || myString.Contains(".") || myString.Contains("?") || myString.Contains("!") || myString.Contains(",") || myString.Contains("(") || myString.Contains(")"))
                    {
                        return "1";
                    }
                    else
                        return "0";

                }
            }
            return "";

        }
        //  words_like_total
        public string containsWordsLikeTotal(Range cell)
        {
            if (cell != null)
            {
                if (cell.Value != null)
                {
                    Object i = cell.Value;
                    string myString = i.ToString();
                    if (myString.Contains("total") || myString.Contains("average") || myString.Contains("avg") || myString.Contains("max") || myString.Contains("min") || myString.Contains("maximum") || myString.Contains("minimum") || myString.Contains("sum") ||
                        myString.Contains("tot") || myString.Contains("ttl") || myString.Contains("tT") || myString.Contains("TL")
                       || myString.Contains("TOTAL") || myString.Contains("TOT") || myString.Contains("TTL") || myString.Contains("tl")
                       || myString.Contains("Total") || myString.Contains("Tot") || myString.Contains("Ttl") || myString.Contains("Tl"))
                    {
                        return "1";
                    }
                    else
                        return "0";

                }
            }
            return "";

        }

        public string containsWordsLikeTable(Range cell)
        {
            if (cell != null)
            {
                if (cell.Value != null)
                {
                    Object i = cell.Value;
                    string myString = i.ToString();
                    if (myString.Contains("table") || myString.Contains("title") || myString.Contains("tab") || myString.Contains("tabl") || myString.Contains("tb") || myString.Contains("Tab")
                       || myString.Contains("TABLE") || myString.Contains("TAB") || myString.Contains("TABL") || myString.Contains("TB") || myString.Contains("Table")
                       || myString.Contains("tbl") || myString.Contains("TBL") || myString.Contains("tlb") || myString.Contains("TLB") || myString.Contains("Tbl")
                       || myString.Contains("Tlb") || myString.Contains("Tb"))
                    {
                        return "1 ";
                    }
                    else
                        return "0";

                }
            }
            return "";

        }
        public string containsColon(Range cell)
        {
            if (cell != null)
            {
                if (cell.Value != null)
                {
                    Object i = cell.Value;
                    string myString = i.ToString();
                    if(myString.Contains(':'))
                    {
                        return "1 ";
                    }
                    else
                        return "0";

                }
            }
            return "";

        }
        public string getCellWidth(Range cell)
        {
            if (cell != null)
            {
                double width = cell.Width;
                return width.ToString();
            }
            return "";

        }


        public string getCellHeight(Range cell)
        {
            if (cell != null)
            {
                double height = cell.Height;
                return height.ToString();
            }
            return "";

        }

        // with consideration of merged cells
        public int getColumnOfCell(Range cell)
        {

            int c = cell.Column;
            if (cell != null)
            {
                //if (cell.MergeCells)
                 //   c = cell.MergeArea.Cells[1, 1].Column;
                return c;


            }
            return 0;

        }
        // with consideration of merged cells
        public int getRowOfCell(Range cell)
        {

            int r = cell.Row;
            if (cell != null)
            {
                //if (cell.MergeCells)
                 //   r = cell.MergeArea.Baby[1, 1].Row;
                return r;


            }
            return 0;
        }

        public int getHorizontalAlignment(Range cell)
        {
            if (cell != null)
            {
                return cell.HorizontalAlignment;
            }
            return 0;
        }

        public int getVerticalAlignment(Range cell)
        {
            if (cell != null)
            {
                return cell.VerticalAlignment;
            }
            return 0;
        }
        public int getFillPattern(Range cell)
        {
            if (cell != null)
            {
                return cell.Interior.Pattern;
            }
            return 0;
        }


        public string isMerged(Range cell) {
            if (cell != null)
            {
                if (cell.MergeCells)
                {
                    return "1";
                }
                else
                    return "0";
            }
            return "9";
        }

        public int getNumberOfCells(Range cell)
        {
           
            if (cell != null)
            {
                if (cell.MergeCells)
                {
                  //  rng = cell.MergeArea;
                    int m = cell.Count;
                    return m;
                }
                return 1;
            }
            return 9;
        }


        public string checkIfStyleMatches(Range cell1, Range cell2)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            Object obj = null;

            if (cell1 != null && cell2 != null)
            {
                obj = cell1.Style;
                //double c1 = getCellFontColor(cell1);
                //double c2 = getCellFontColor(cell2);
                double s1 = getFontSize(cell1);
                double s2 = getFontSize(cell2);
                string b1 = isBold(cell1);
                string b2 = isBold(cell2);
                string n1 = getFontName(cell1);
                string n2 = getFontName(cell2);
                string i1 = isItalic(cell1);
                string i2 = isItalic(cell2);
                string u1 = getUnderlineType(cell1);
                string u2 = getUnderlineType(cell2);
                string o1 = getOrientation(cell1);
                string o2 = getOrientation(cell2);
                int h1 = getHorizontalAlignment(cell1);
                int h2 = getHorizontalAlignment(cell2);
                int v1 = getVerticalAlignment(cell1);
                int v2 = getVerticalAlignment(cell2);
                double f1 = getCellFillColor(cell1);
                double f2 = getCellFillColor(cell2);

                if ( s1 == s2
                    && n1 == n2
                    && b1 == b2
                    && i1 == i2
                    && u1 == u2
                    && o1 == o2
                    && h1 == h2
                    && v1 == v2
                    && f1 == f2)
                    return "1";
                else
                    return "0";
            }
            else return "9";
        }
    

        public Range getTopNeighbor(int r, int c)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            Range cell;
            if (r - 1 > 0)
            {

                if (currentSheet.Cells[r - 1, c].Value2 != null)
                {
                    cell = currentSheet.Cells[r - 1, c];
                    return cell;
                }

            }
            else
                return null;
            return null;
           
        }

        public Range getLeftNeighbor(int r, int c)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
          Range cell;

            if (c - 1 > 0)
            {
                if (currentSheet.Cells[r, c - 1].Value2 != null)
                {
                   cell = currentSheet.Cells[r, c - 1];
                    return cell;
                }

            }
            else
                return null;
            return null;

        }

        public Range getRightNeighbor(int r, int c)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            Range cell;

            if (currentSheet.Cells[r, c + 1].Value2 != null)
            {
                cell = currentSheet.Cells[r, c + 1];
                return cell;
            }
            else
                return null;
        }

        public Range getBottomNeighbor(int r, int c)
        {

            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            Range cell;

            if (currentSheet.Cells[r + 1, c].Value2 != null)
            {
                cell = currentSheet.Cells[r + 1, c];
                return cell;
            }
            else
                return null;
        }

        //  IS_AGGREGATION_FORMULA
        public string checkIfCellContainsAggregationFormula(Range rng)
        {
            if (rng.Value != null)
            {
                if (rng.HasFormula)
                {
                    string formula = rng.Formula.ToString();
                    if (formula.Contains("MAX") || formula.Contains("MIN") || formula.Contains("AVERAGE") || formula.Contains("AVERAGEA") || formula.Contains("COUNT")
                        || formula.Contains("COUNTA") || formula.Contains("MAXX") || formula.Contains("MINX") || formula.Contains("SUM")
                        || formula.Contains("SUMME") || formula.Contains("PRODUCT") || formula.Contains("SUBTOTAL") || formula.Contains("STDEV") || formula.Contains("STDEV.P")
                        || formula.Contains("STDEV.S") || formula.Contains("STDEVA") || formula.Contains("STDEVP") || formula.Contains("STDEVPA")
                        || formula.Contains("VAR") || formula.Contains("VAR.P") || formula.Contains("VARA") || formula.Contains("VARP")
                        || formula.Contains("AGGREGATE") || formula.Contains("MEDIAN") || formula.Contains("MODE.MULT") || formula.Contains("MODE.SNGL")
                        || formula.Contains("LARGE") || formula.Contains("MODE") || formula.Contains("SMALL") || formula.Contains("MODE.SNGL")
                        || formula.Contains("PERCENTILE") || formula.Contains("PERCENTILE.EXC") || formula.Contains("QUARTILE.EXC") || formula.Contains("QUARTILE.INC")
                        || formula.Contains("QUARTILE"))
                        return "1";
                    else
                        return "0";
                }
            }
            return "9";

        }

        public string checkIfStringContainsAggregationFormula(String str)
        {
            if (str != null)
            {

                if (str.Contains("MAX") || str.Contains("MIN") || str.Contains("AVERAGE") || str.Contains("AVERAGEA") || str.Contains("COUNT")
                    || str.Contains("COUNTA") || str.Contains("MAXX") || str.Contains("MINX") || str.Contains("SUM")
                    || str.Contains("SUMME") || str.Contains("PRODUCT") || str.Contains("SUBTOTAL") || str.Contains("STDEV") || str.Contains("STDEV.P")
                    || str.Contains("STDEV.S") || str.Contains("STDEVA") || str.Contains("STDEVP") || str.Contains("STDEVPA")
                    || str.Contains("VAR") || str.Contains("VAR.P") || str.Contains("VARA") || str.Contains("VARP")
                    || str.Contains("AGGREGATE") || str.Contains("MEDIAN") || str.Contains("MODE.MULT") || str.Contains("MODE.SNGL")
                    || str.Contains("LARGE") || str.Contains("MODE") || str.Contains("SMALL") || str.Contains("MODE.SNGL")
                    || str.Contains("PERCENTILE") || str.Contains("PERCENTILE.EXC") || str.Contains("QUARTILE.EXC") || str.Contains("QUARTILE.INC")
                    || str.Contains("QUARTILE"))
                    return "1";
                else
                    return "0";
            }

            return "9";

        }


        // vo cell sodrzinata aggregationsformula da dobie svoj range  
        // dadena formula dobiva lista so sie svoi sodrzani range-ovi
        // ima skrieni elementi samo koga ima dve tocki!!! 
        // funkcionira super ! 
        public List<Range> findReferencedCellsInAggregationFormula(Range cell)
        {


            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();


            object obj = cell.Formula;
            string cellFormula = obj.ToString();
            string onTheLeft = ""; //primer
            string onTheRight = ""; // primer
            List<Range> lst = new List<Range>();


            var count = cellFormula.Count(x => x == ':');



            while (count > 0)
            {


                for (int i = 0; i < cellFormula.ToCharArray().Length; i++)
                {
                    if (cellFormula.ToCharArray().ElementAt(i) == ':')
                    {
                        onTheLeft = String.Concat(cellFormula.ToCharArray().ElementAt(i - 2), cellFormula.ToCharArray().ElementAt(i - 1));
                        onTheRight = String.Concat(cellFormula.ToCharArray().ElementAt(i + 1), cellFormula.ToCharArray().ElementAt(i + 2));

                        count--;

                       

                        lst.Add(currentSheet.Range[onTheLeft, onTheRight]);
                    

                    }


                }
            }

            return lst;


        }

        //ref_val_type
        // se zema krajniot rezultat na celata formula 
        // funkcionira super 
        public string getRefValValue(Range cell)
        {

            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            int numericTypes = 0;
            int stringTypes = 0;
            int booleanTypes = 0;
            string celladdress = getLabel(cell.Address.ToString());
            Object valueOfFormula = null;
            string valOfFormula = "";

            Range rng = currentSheet.UsedRange;

            foreach (Range r in rng)
            {
                if (r.HasFormula)
                {

                   
                    object obj = r.Formula;
                    string cFormula = obj.ToString();

                    List<Range> lst = new List<Range>();

                  

                    // !!! ovde fali toa kaj so vo lst = findReferencedCellsInAggregationFormula(r), POPRAVI
                    lst = findReferencedCellsInAggregationFormula(r);


                  
/*
                    for (int i = 0; i < lst.Count; i++) {
                        Debug.WriteLine("Element e  " + lst.ElementAt(i));
                    } */

                    if (cFormula.Contains(celladdress) || checkIfElementIsInList(lst, cell) == "1")
                    {
                        valueOfFormula = r.Value;
                        valOfFormula = valueOfFormula.ToString();

                     
                        string result = checkTypeOfValue(valueOfFormula);

                      
                        if (result == "NUMERIC")
                            numericTypes++;
                        if (result == "STRING")
                            stringTypes++;
                        if (result == "BOOLEAN")
                            booleanTypes++;

                    }

                }

            }
            int[] values;
            values = new int[5];
            values[0] = numericTypes;
            values[1] = stringTypes;
            values[2] = booleanTypes;

         


            if (values[0] == 0 && values[1] == 0 && values[2] == 0)
                return "9";

            if (numericTypes == stringTypes && stringTypes > booleanTypes)
                return "STRING";
            if (numericTypes == booleanTypes && booleanTypes > stringTypes)
                return "BOOLEAN";

            if (values.Max() == values[0])
                return "NUMERIC";
            if (values.Max() == values[1])
                return "STRING";
            if (values.Max() == values[2])
                return "BOOLEAN";

          //  Debug.WriteLine("Max e " + values.Max());

            return "";

        }

        public string checkTypeOfValue(Object s)
        {

            Type t = s.GetType();


            if (t.Equals(typeof(int)))
                return contentType.NUMERIC.ToString();
            if (t.Equals(typeof(double)))
                return contentType.NUMERIC.ToString();
            if (t.Equals(typeof(float)))
                return contentType.NUMERIC.ToString();
            if (t.Equals(typeof(String)))
                return contentType.STRING.ToString();
            if (t.Equals(typeof(Boolean)))
                return contentType.BOOLEAN.ToString();
            if (t.Equals(typeof(DateTime)))
                return contentType.DATE.ToString();



            //return contentType.STRING.ToString();
            return "9";
        }

        public string checkIfElementIsInList(List<Range> lst, Range cell)
        {

            foreach (Range r in lst)
            {
                if (Globals.ThisAddIn.Application.Intersect(r, cell) != null)
                {

                    if (Globals.ThisAddIn.Application.Intersect(r, cell).Address == cell.Address)
                        return "1";
                }

            }
            return "0";
        }

        //REF_IS_AGGREGATION_FORMULA
        public string checkIfRefIsAggregationFormula(Range cell)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            string celladdress = getLabel(cell.Address.ToString());
            string firstOfDoublePointFormula = "";

            List<Range> myList = new List<Range>();



            Range rng = currentSheet.UsedRange;

            foreach (Range r in rng)
            {
                if (r.HasFormula)
                {
                    object obj = r.Formula;
                    string rngFormula = obj.ToString();


              

                    int position = 0;
   
                 
                   

                    if (rngFormula.Contains(getLabel(cell.Address)) || checkIfElementIsInList(findReferencedCellsInAggregationFormula(currentSheet.Range[getLabel(r.Address)]), cell) == "1") // tuka da se dodade mislam :D 
                    {
                        List<Range> list = new List<Range>();
                        if (checkIfElementIsInList(findReferencedCellsInAggregationFormula(currentSheet.Range[getLabel(r.Address)]), cell) == "1")
                        {
                           Range ro = rangeContainingCell(findReferencedCellsInAggregationFormula(currentSheet.Range[getLabel(r.Address)]), cell);

                            foreach (Range c in ro.Cells) {
                                list.Add(c);
                            }

                            firstOfDoublePointFormula = list.ElementAt(0).Address.ToString();
                          
                        }


                        for (int i = 0; i < rngFormula.Count<char>() - 1; i++)
                        {

                            string substring = String.Concat(rngFormula[i], rngFormula[i + 1]);  


                            if (substring == getLabel(cell.Address) || substring == getLabel(firstOfDoublePointFormula))
                            {
                                position = i;

                                break;
                            }
                        
                       }

                        int bracketsCounter = 0;

                        string function = "";
                        List<char> lst = rngFormula.ToList<char>();

                        int indexOfFirstBracket = 0;

                        for (int i = position; i > 0; i--)
                        {
                          
                            if (rngFormula.ToCharArray().ElementAt(i) == '(')
                            {


                                indexOfFirstBracket = i;


                                break;
                            }
                           
                        }

                        for (int i = position; i > 0; i--)
                        {


                            if (rngFormula.ToCharArray().ElementAt(i) == '(')
                            {
                                bracketsCounter++;

                            }

                            if (rngFormula.ToCharArray().ElementAt(i) == ')')
                                bracketsCounter--;
                        }


                        String set = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.";
                        int positionOfNonLetterOrDotCharacter = 0;

                        if (bracketsCounter > 0) 
                        {
                            
                            for (int i = indexOfFirstBracket - 1; i > 0; i--)
                            {
                                if (!set.Contains(rngFormula.ToCharArray().ElementAt(i)))
                                {
                                    positionOfNonLetterOrDotCharacter = i;

                                    break;

                                }
                            }
                        }


                   


                        for (int j = positionOfNonLetterOrDotCharacter; j < indexOfFirstBracket; j++)
                        {
                            function = string.Concat(function, rngFormula.ToCharArray().ElementAt(j));
                        }


                     


                        if (checkIfStringContainsAggregationFormula(function) == "1")
                            return "1";
                        else
                            return "0";

                    }



                }


            }
            return "9";

        }


        // FORMULA_VAL_TYPE
        public string getReturnTypeOfFormula(Range cell)
        {

            if (cell.HasFormula) {
             return checkFormulaType(cell);
            }
            else return "9"; // not applicable

        }


        public Range rangeContainingCell(List<Range> lst, Range cell)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            Range rng = currentSheet.Range["A1"]; 

            for (int i = 0; i < lst.Count; i++)
            {

                if (Globals.ThisAddIn.Application.Intersect(lst.ElementAt(i), cell) != null)
                {

                    if (Globals.ThisAddIn.Application.Intersect(lst.ElementAt(i), cell).Address == cell.Address)
                        rng = lst.ElementAt(i);
                }
            }
            return rng;
        }

        //Classify selection  (should be used when we want to classify only a selected part of the table) 
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            Range selectedRange = Globals.ThisAddIn.Application.Selection;
            Debug.WriteLine(" Selected range e golem: " + selectedRange.Count);
            IterateCells(currentSheet, selectedRange);

            RunAsync(getUrl1(), getUrl2()).Wait();
        }


        // show data label 
        private async void checkBox1_Click(object sender, RibbonControlEventArgs e)
        {
            if (checkIfUnchecked1() == false)
                colorRangeWhite();
            else
            await showDataLabel("http://127.0.0.1:5000/classify/xcells/api/v0.2/tasks/run/");
         

        }

        //header
        private async void checkBox2_Click(object sender, RibbonControlEventArgs e)
        {
            if (checkIfUnchecked2() == false)
                colorRangeWhite();
            else
                await  showHeaderLabel("http://127.0.0.1:5000/classify/xcells/api/v0.2/tasks/run/");

        }

        //derived
        private async void checkBox3_Click(object sender, RibbonControlEventArgs e)
        {
            if (checkIfUnchecked3() == false)
                colorRangeWhite();
            else
                await  showDerivedLabel("http://127.0.0.1:5000/classify/xcells/api/v0.2/tasks/run/");
        }
  
        //metadata
        private async void checkBox4_Click(object sender, RibbonControlEventArgs e)
        {
            if (checkIfUnchecked4() == false)
                colorRangeWhite();
            else
                await  showMetadataLabel("http://127.0.0.1:5000/classify/xcells/api/v0.2/tasks/run/");
        }

        //attributes 
        private void checkBox5_Click(object sender, RibbonControlEventArgs e)
        {
            if (checkIfUnchecked5() == false)
                colorRangeWhite();
            else
                showAttributesLabel("http://127.0.0.1:5000/classify/xcells/api/v0.2/tasks/run/");
        }

        public Boolean checkIfUnchecked1(){
            return checkBox1.Checked;
        }
        public Boolean checkIfUnchecked2()
        {
            return checkBox2.Checked;
        }
        public Boolean checkIfUnchecked3()
        {
            return checkBox3.Checked;
        }
        public Boolean checkIfUnchecked4()
        {
            return checkBox4.Checked;
        }
        public Boolean checkIfUnchecked5()
        {
            return checkBox5.Checked;
        }

        public async Task showDataLabel(string path) {

            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();

            int nextId = await getNextId("http://127.0.0.1:5000/classify/xcells/api/v0.2/tasks");

            RootObject product = null;
            HttpResponseMessage response = await client.GetAsync(path + nextId);
            var responseBody = response.Content.ReadAsStringAsync().Result;

            RootObjectDes data = JsonConvert.DeserializeObject<RootObjectDes>(responseBody);


            Debug.WriteLine("Responsebody e " + responseBody);

            if (response.IsSuccessStatusCode)
            {
                product = await response.Content.ReadAsAsync<RootObject>();
            }

            foreach (var item in data.results)
            {
                foreach (var cell in item.cells)
                {
                   Range rng = currentSheet.Range[cell.address.ToString()];


                    System.Drawing.Color color1 = System.Drawing.Color.FromArgb(204, 255, 204); // green

                    if (cell.predicted == "data")
                        rng.Interior.Color = System.Drawing.ColorTranslator.ToOle(color1);
                    else continue;

                }
            }
        }

        // show header label 

        public async Task showHeaderLabel(string path)
        {

            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            int nextId = await getNextId("http://127.0.0.1:5000/classify/xcells/api/v0.2/tasks");
            RootObject product = null;
            HttpResponseMessage response = await client.GetAsync(path + nextId);

            var responseBody = response.Content.ReadAsStringAsync().Result;

            RootObjectDes data = JsonConvert.DeserializeObject<RootObjectDes>(responseBody);


            Debug.WriteLine("Responsebody e " + responseBody);

            if (response.IsSuccessStatusCode)
            {
                product = await response.Content.ReadAsAsync<RootObject>();
            }

            foreach (var item in data.results)
            {

                foreach (var cell in item.cells)
                {

                    Range rng = currentSheet.Range[cell.address.ToString()];

                    Debug.WriteLine(" Adresa na cell e  " + cell.address);

                    System.Drawing.Color color2 = System.Drawing.Color.FromArgb(180, 205, 205); //silver

                    if (cell.predicted == "header")
                        rng.Interior.Color = System.Drawing.ColorTranslator.ToOle(color2);
                    else continue;

                }
            }
        }

        // show derived label 
        public async Task showDerivedLabel(string path)
        {

            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            int nextId = await getNextId("http://127.0.0.1:5000/classify/xcells/api/v0.2/tasks");
            RootObject product = null;
            HttpResponseMessage response = await client.GetAsync(path + nextId);

            var responseBody = response.Content.ReadAsStringAsync().Result;

            RootObjectDes data = JsonConvert.DeserializeObject<RootObjectDes>(responseBody);


            Debug.WriteLine("Responsebody e " + responseBody);

            if (response.IsSuccessStatusCode)
            {
                product = await response.Content.ReadAsAsync<RootObject>();
            }

            foreach (var item in data.results)
            {

                foreach (var cell in item.cells)
                {

                    Range rng = currentSheet.Range[cell.address.ToString()];

                    System.Drawing.Color color3 = System.Drawing.Color.FromArgb(254, System.Drawing.Color.LightCoral); //coral

                    if (cell.predicted == "derived")
                        rng.Interior.Color = System.Drawing.ColorTranslator.ToOle(color3);
                    else continue;

                }
            }
        }

        // show metadata label 
        public async Task showMetadataLabel(string path)
        {

            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            int nextId = await getNextId("http://127.0.0.1:5000/classify/xcells/api/v0.2/tasks");
            RootObject product = null;
            HttpResponseMessage response = await client.GetAsync(path + nextId);

            var responseBody = response.Content.ReadAsStringAsync().Result;

            RootObjectDes data = JsonConvert.DeserializeObject<RootObjectDes>(responseBody);


            Debug.WriteLine("Responsebody e " + responseBody);

            if (response.IsSuccessStatusCode)
            {
                product = await response.Content.ReadAsAsync<RootObject>();
            }

            foreach (var item in data.results)
            {

                foreach (var cell in item.cells)
                {

                    Range rng = currentSheet.Range[cell.address.ToString()];

                
              
                    System.Drawing.Color color4 = System.Drawing.Color.FromArgb(255, 255, 179); //yellow
                   
                    if (cell.predicted == "metadata")
                        rng.Interior.Color = System.Drawing.ColorTranslator.ToOle(color4);
                    else continue;

                }
            }
        }

        // show attributes label 
        public async Task showAttributesLabel(string path)
        {

            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            int nextId = await getNextId("http://127.0.0.1:5000/classify/xcells/api/v0.2/tasks");
            RootObject product = null;
            HttpResponseMessage response = await client.GetAsync(path + nextId);

            var responseBody = response.Content.ReadAsStringAsync().Result;

            RootObjectDes data = JsonConvert.DeserializeObject<RootObjectDes>(responseBody);


            Debug.WriteLine("Responsebody e " + responseBody);

            if (response.IsSuccessStatusCode)
            {
                product = await response.Content.ReadAsAsync<RootObject>();
            }

            foreach (var item in data.results)
            {

                foreach (var cell in item.cells)
                {

                    Range rng = currentSheet.Range[cell.address.ToString()];

                  

                    System.Drawing.Color color5 = System.Drawing.Color.FromArgb(213, 186, 219); // violet

                    if (cell.predicted  == "attributes")
                        rng.Interior.Color = System.Drawing.ColorTranslator.ToOle(color5);
                    else continue;

                }
            }
        }

        // color white (Start again)
        private void button2_Click_1(object sender, RibbonControlEventArgs e)
        {
            colorRangeWhite();
        }
       
        public static void colorRangeWhite() {
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();
            Range usedRange = currentSheet.UsedRange;

            foreach (Range cell in usedRange)
            {
                cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            }
        }

    }
}

  
    

    





 