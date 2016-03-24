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
    public class MiscUtility
    {
        public static void LogHistory(string text)
        {
            string logfilename = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "History.log");
            SaveFile(logfilename, string.Format("[{0}] - {1}", DateTime.Now.ToString(), text));
        }

        public static void LogError(string text)
        {
            string logfilename = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Error.log");
            SaveFile(logfilename, string.Format("[{0}] - {1}", DateTime.Now.ToString(), text));
        }

        public static void InsertNewColumn(ref System.Data.DataTable datatablename, string columnname, string columntype, int columnindex)
        {
            DataColumn dc = new DataColumn();
            dc.ColumnName = columnname;
            dc.DataType = System.Type.GetType(columntype);
            datatablename.Columns.Add(dc);

            dc.SetOrdinal(columnindex - 1);
        }

        public static void SaveFile(string filename, string text)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(filename, true))
                {
                    writer.WriteLine(text);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Saving the file - {0} failed: {1}", filename, ex.Message));
            }
        }

        private string LoadTextFile(string path)
        {
            string text = null;

            try
            {
                using (StreamReader reader = new FileInfo(path).OpenText())
                {
                    text = reader.ReadToEnd();
                }

                return text;
            }
            catch (FileNotFoundException ex)
            {
                throw new FileNotFoundException(
                    string.Format("The file {0} cannot be found: {1}", path, ex.Message));
            }
            catch (FileLoadException ex)
            {
                throw new FileLoadException(
                    string.Format("Loading the file {0} failed: {1}", path, ex.Message));
            }
            catch
            {
                throw;
            }
        }        
    }
}
