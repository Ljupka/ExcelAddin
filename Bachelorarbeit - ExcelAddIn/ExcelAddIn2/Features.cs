using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn2
{ 
   public class Features
    {
        public string value { get; set; }
        // font features
        public string FONT_COLOR_DEFAULT { get; set; }
        public string FONT_SIZE { get; set; }
        public string IS_BOLD { get; set; }
        public string IS_ITALIC { get; set; }
        public string IS_STRIKE_OUT { get; set; }
        public string UNDERLINE_TYPE { get; set; }
        public string IS_UNDERLINED { get; set; }
        public string OFFSET_TYPE { get; set; }

        // spatial features 
        public string ROW_NUM { get; set; }
        public string COLUMN_NUM { get; set; }
        public string NUMBER_OF_NEIGHBORS { get; set; }
        public string MATCHES_TOP_STYLE { get; set; }
        public string MATCHES_BOTTOM_STYLE { get; set; }
        public string MATCHES_LEFT_STYLE { get; set; }
        public string MATCHES_RIGHT_STYLE { get; set; }


        public string MATCHES_TOP_TYPE { get; set; }
        public string MATCHES_BOTTOM_TYPE { get; set; }
        public string MATCHES_LEFT_TYPE { get; set; }
        public string MATCHES_RIGHT_TYPE { get; set; }

        public string TOP_NEIGHBOR_TYPE { get; set; }
        public string BOTTOM_NEIGHBOR_TYPE { get; set; }
        public string LEFT_NEIGHBOR_TYPE { get; set; }
        public string RIGHT_NEIGHBOR_TYPE { get; set; }


        //content features 
        public string LENGTH { get; set; }
        public string NUM_OF_TOKENS { get; set; }
        public string LEADING_SPACES { get; set; }
        public string IS_NUMERIC { get; set; }
        public string IS_FORMULA { get; set; }
        public string STARTS_WITH_NUMBER { get; set; }
        public string STARTS_WITH_SPECIAL { get; set; }
        public string IS_CAPITALIZED { get; set; }
        public string IS_UPPER_CASE { get; set; }
        public string IS_ALPHABETIC { get; set; }
        public string IS_ALPHANUMERIC { get; set; }
        public string CONTAINS_SPECIAL_CHARS { get; set; }
        public string CONTAINS_PUNCTUATIONS { get; set; }
        public string CONTAINS_COLON { get; set; }
        public string WORDS_LIKE_TOTAL { get; set; }
        public string WORDS_LIKE_TABLE { get; set; }
        public string IN_YEAR_RANGE { get; set; }
        

        //cell style features 

        public string HORIZONTAL_ALIGNMENT { get; set; }
        public string VERTICAL_ALIGNMENT { get; set; }
        public string FILL_COLOR_DEFAULT { get; set; }
        public string FILL_PATTERN { get; set; }
        public string ORIENTATION { get; set; }
        public string CONTROL { get; set; }
        public string IS_MERGED { get; set; }
        public string NUMBER_OF_CELLS { get; set; }
        public string BORDER_TOP_TYPE { get; set; }
        public string BORDER_BOTTOM_TYPE { get; set; }
        public string BORDER_LEFT_TYPE { get; set; }
        public string BORDER_RIGHT_TYPE { get; set; }
        public string NUM_OF_BORDERS { get; set; }

        public string INDENTATIONS { get; set; }

        public string BORDER_TOP_THICKNESS { get; set; }
        public string BORDER_BOTTOM_THICKNESS { get; set; }
        public string BORDER_LEFT_THICKNESS { get; set; }
        public string BORDER_RIGHT_THICKNESS { get; set; }


        public string FONT_COLOR { get; set; }
        public string FONT_NAME { get; set; }
      

        // referencing features 
        public string FORMULA_VAL_TYPE { get; set; }
        public string IS_AGGREGATION_FORMULA { get; set; }
        public string REF_VAL_TYPE { get; set; }
        public string REF_IS_AGGREGATION_FORMULA { get; set; }

        // extras

        public string WIDTH { get; set; }
        public string HEIGHT { get; set; }




       
        // cell description 
        //public string address { get; set; }
        //public string label { get; set; }



    }
}
 