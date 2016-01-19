using System;
using System.Data;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Macro.Common.API.Excel
{
    public class Excel
    {
        public Excel()
        {

        }

        public System.Data.DataTable Read(string path)
        {
            Application _excelApp = new Application();

            //open the workbook
            Workbook workbook = _excelApp.Workbooks.Open(path,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            //select the first sheet        
            Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

            //find the used range in worksheet
            Range excelRange = worksheet.UsedRange;

            //get an object array of all of the cells in the worksheet (their values)
            object[,] valueArray = (object[,])excelRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);

            System.Data.DataTable records = new System.Data.DataTable();

            //access the cells
            for (int row = 1; row <= worksheet.UsedRange.Rows.Count; ++row)
            {
                DataRow dataRow = records.NewRow();
                for (int col = 1; col <= worksheet.UsedRange.Columns.Count; ++col)
                {
                    if (row == 1)
                        records.Columns.Add(valueArray[row, col].ToString().ToUpper().Trim(), valueArray[row, col].GetType());
                    else
                        dataRow[col - 1] = valueArray[row, col];
                }

                if (row > 1)
                    records.Rows.Add(dataRow);
            }

            //clean up stuffs
            workbook.Close(false, Type.Missing, Type.Missing);
            Marshal.ReleaseComObject(workbook);

            _excelApp.Quit();
            Marshal.FinalReleaseComObject(_excelApp);

            return records;
        }
    }
}
