using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using XlBorderWeight = Microsoft.Office.Interop.Excel.XlBorderWeight;

namespace DAO_3PL_Report_Tool
{
    public class ExcelFileUtility
    {
        public static void ExportDataIntoCSVFile(string filename, System.Data.DataTable datatable)
        {
            if (filename.Length != 0)
            {
                FileStream filestream = null;
                StreamWriter streamwriter = null;
                string stringline = string.Empty;

                try
                {
                    filestream = new FileStream(filename, FileMode.Append, FileAccess.Write);
                    streamwriter = new StreamWriter(filestream, System.Text.Encoding.Unicode);

                    for (int i = 0; i < datatable.Columns.Count; i++)
                    {
                        stringline = stringline + datatable.Columns[i].ColumnName.ToString() + Convert.ToChar(9);

                    }

                    streamwriter.WriteLine(stringline);
                    stringline = "";

                    for (int i = 0; i < datatable.Rows.Count; i++)
                    {
                        //stringline = stringline + (i + 1) + Convert.ToChar(9);
                        for (int j = 0; j < datatable.Columns.Count; j++)
                        {
                            stringline = stringline + datatable.Rows[i][j].ToString() + Convert.ToChar(9);
                        }

                        streamwriter.WriteLine(stringline);
                        stringline = "";
                    }

                    streamwriter.Close();
                    filestream.Close();
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public static void SaveExcelFile(string filename, DataTable datatable, bool iswithline)
        {
            Microsoft.Office.Interop.Excel.Application excel = null;
            Workbook workbook = null;
            //Worksheet worksheet = null;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();

                if (excel == null)
                    throw new Exception("There is not an Excel application on your computer!");

                excel.Application.Workbooks.Add(true);
                excel.Visible = false;
                excel.DisplayAlerts = false;

                workbook = excel.Workbooks.Add();
                Worksheet worksheet = (Worksheet)workbook.ActiveSheet;

                // Write column name into Excel file
                int colIndex = 0;
                foreach (DataColumn col in datatable.Columns)
                {
                    colIndex++;
                    excel.Cells[1, colIndex] = col.ColumnName;
                }

                Range range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, colIndex]];
                if (iswithline)
                {
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlThin;
                }

                // Write row data into Excel file
                int rowcount = datatable.Rows.Count;
                int colcount = datatable.Columns.Count;

                if (rowcount != 0 && colcount != 0)
                {
                    var dataarray = new object[rowcount, colcount];

                    for (int indey = 0; indey < rowcount; indey++)
                    {
                        for (int indez = 0; indez < colcount; indez++)
                        {
                            dataarray[indey, indez] = datatable.Rows[indey][indez];
                        }
                    }

                    range = worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[rowcount + 1, colcount]];
                    range.Value = dataarray;
                }

                if (iswithline)
                {
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlThin;
                }

                worksheet.Cells.EntireColumn.AutoFit();
                SetTitusClassification(ref workbook);
                workbook.SaveAs(filename);
            }

            catch (Exception ex)
            {
                MiscUtility.LogError(
                    string.Format("Message:{0}, Source:{1}, StackTrack:{2}", ex.InnerException.Message, ex.Source, ex.StackTrace));
            }

            finally
            {
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);

                if (workbook != null)
                {
                    workbook.Close();
                    Marshal.ReleaseComObject(workbook);
                    excel = null;
                }

                if (excel != null)
                {
                    excel.Quit();
                    Marshal.ReleaseComObject(excel);
                    workbook = null;
                }

                GC.Collect();
            }
        }

        // Add some properties for Titus Classification into Excel file.
        public static void SetTitusClassification(ref Workbook workBook)
        {
            SetDocumentProperty(ref workBook, "DellClassification", "Internal Use");
            SetDocumentProperty(ref workBook, "TitusReset", "Reset");
        }

        // Setup a customer property for excel file.
        public static void SetDocumentProperty(ref Workbook workBook,
            string propertyName, string propertyValue)
        {
            dynamic oDocCustomProps = workBook.CustomDocumentProperties;
            Type typeDocCustomProps = oDocCustomProps.GetType();

            dynamic[] oArgs = { propertyName, false, MsoDocProperties.msoPropertyTypeString, propertyValue };

            typeDocCustomProps.InvokeMember("Add", BindingFlags.Default | BindingFlags.InvokeMethod, null,
                oDocCustomProps, oArgs);
        }

        public static dynamic GetDocumentProperty(ref Workbook workBook,
            string propertyName, MsoDocProperties type)
        {
            dynamic returnVal = null;

            dynamic oDocCustomProps = workBook.CustomDocumentProperties;
            Type typeDocCustomProps = oDocCustomProps.GetType();

            dynamic returned = typeDocCustomProps.InvokeMember("Item",
                BindingFlags.Default |
                BindingFlags.GetProperty, null,
                oDocCustomProps, new object[] { propertyName });

            Type typeDocAuthorProp = returned.GetType();
            returnVal = typeDocAuthorProp.InvokeMember("Value",
                BindingFlags.Default |
                BindingFlags.GetProperty,
                null, returned,
                new object[] { }).ToString();

            return returnVal;
        }
    }
}
