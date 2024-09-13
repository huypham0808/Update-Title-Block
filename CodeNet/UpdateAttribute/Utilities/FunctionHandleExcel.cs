using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace UpdateAttribute.Utilities
{
    public class FunctionHandleExcel
    {
        public static int ExcelColumnToIndex(string column)
        {
            int columnIndex = 0;
            int columnBase = 'Z' - 'A' + 1;

            foreach (char c in column)
            {
                columnIndex = columnIndex * columnBase + c - 'A' + 1;
            }

            return columnIndex;
        }
        public static string IndexToExcelColumn(int index)
        {
            string column = "";
            int columnBase = 'Z' - 'A' + 1;

            while (index > 0)
            {
                int remainder = (index - 1) % columnBase;
                column = (char)('A' + remainder) + column;
                index = (index - 1) / columnBase;
            }
            return column;
        }
    }
}
