using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Drawing;

namespace Macro.Common.API.Excel
{
    public static class ExcelListExtension
    {
        public static void ToExcel<T>(this List<T> list, string pathToSave)
        {
            #region Declarations

            if (string.IsNullOrEmpty(pathToSave))
            {
                throw new Exception("Invalid file path.");
            }
            else if (pathToSave.ToLower().Contains("") == false)
            {
                throw new Exception("Invalid file path.");
            }

            if (list == null)
            {
                throw new Exception("No data to export.");
            }

            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Microsoft.Office.Interop.Excel.Workbooks books = null;
            Microsoft.Office.Interop.Excel._Workbook book = null;
            Microsoft.Office.Interop.Excel.Sheets sheets = null;
            Microsoft.Office.Interop.Excel._Worksheet sheet = null;
            Microsoft.Office.Interop.Excel.Range range = null;
            Microsoft.Office.Interop.Excel.Font font = null;
            // Optional argument variable
            object optionalValue = Missing.Value;

            string strHeaderStart = "A2";
            string strDataStart = "A3";
            #endregion

            #region Processing

            try
            {
                #region Init Excel app.

                excelApp = new Microsoft.Office.Interop.Excel.Application();
                books = excelApp.Workbooks;
                book = books.Add(optionalValue);
                sheets = book.Worksheets;
                sheet = (Microsoft.Office.Interop.Excel._Worksheet)(sheets.get_Item(1));

                #endregion

                #region Creating Header

                Dictionary<string, string> objHeaders = new Dictionary<string, string>();

                PropertyInfo[] headerInfo = typeof(T).GetProperties();

                foreach (var property in headerInfo)
                {
                    var attribute = property.GetCustomAttributes(typeof(DisplayNameAttribute), false)
                                            .Cast<DisplayNameAttribute>().FirstOrDefault();

                    objHeaders.Add(property.Name, attribute == null ?
                                        property.Name : attribute.DisplayName);
                }

                range = sheet.get_Range(strHeaderStart, optionalValue);
                range = range.get_Resize(1, objHeaders.Count);

                range.set_Value(optionalValue, objHeaders.Values.ToArray());
                range.BorderAround(Type.Missing, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);

                font = range.Font;
                font.Bold = true;
                range.Interior.Color = Color.LightGray.ToArgb();

                #endregion

                #region Writing data to cell


                int count = list.Count;
                object[,] objData = new object[count, objHeaders.Count];

                for (int j = 0; j < count; j++)
                {
                    var item = list[j];
                    int i = 0;
                    int k = 0;
                    foreach (KeyValuePair<string, string> entry in objHeaders)
                    {
                        var y = typeof(T).InvokeMember(entry.Key.ToString(), BindingFlags.GetProperty, null, item, null);

                        if (k == 0)
                        {
                            objData[j, i++] = "";
                            k++;
                            continue;
                        }

                        if (y != null && y is decimal)
                        {
                            decimal d = (decimal)y;
                            objData[j, i++] = (y == null) ? "" : d.ToString("#,#.00#");

                            continue;
                        }
                        
                        objData[j, i++] = (y == null) ? "" : y.ToString();

                    }
                }


                range = sheet.get_Range(strDataStart, optionalValue);
                range = range.get_Resize(count, objHeaders.Count);

                range.set_Value(optionalValue, objData);
                range.BorderAround(Type.Missing, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Type.Missing);

                range = sheet.get_Range(strHeaderStart, optionalValue);
                range = range.get_Resize(count + 1, objHeaders.Count);
                range.Columns.AutoFit();

                #endregion

                #region Saving data and Opening Excel file.


                if (string.IsNullOrEmpty(pathToSave) == false)
                    book.SaveAs(pathToSave);

                excelApp.Visible = true;

                #endregion

                #region Release objects

                try
                {
                    if (sheet != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                    sheet = null;

                    if (sheets != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(sheets);
                    sheets = null;

                    if (book != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(book);
                    book = null;

                    if (books != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(books);
                    books = null;

                    if (excelApp != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                    excelApp = null;
                }
                catch (Exception ex)
                {
                    sheet = null;
                    sheets = null;
                    book = null;
                    books = null;
                    excelApp = null;
                }
                finally
                {
                    GC.Collect();
                }

                #endregion
            }
            catch (Exception ex)
            {
                throw ex;
            }

            #endregion
        }
    }
}
