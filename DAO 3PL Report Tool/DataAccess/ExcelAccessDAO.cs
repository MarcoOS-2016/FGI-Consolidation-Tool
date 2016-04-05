using System;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DAO_3PL_Report_Tool
{
    public class ExcelAccessDAO : ExcelFileAccess
    {
        public ExcelAccessDAO()
        {
        }

        public ExcelAccessDAO(string filename)
            : base(filename)
        {
            try
            {
                if (base.connection.State == System.Data.ConnectionState.Closed)
                {
                    base.connection.Open();
                }
            }
            catch
            {
                throw;
            }
        }

        public ExcelAccessDAO(string filename, bool isfield)
            : base(filename, isfield)
        {
            try
            {
                if (base.connection.State == System.Data.ConnectionState.Closed)
                {
                    base.connection.Open();
                }
            }
            catch
            {
                throw;
            }
        }

        public DataTable GetExcelSheetName()
        {
            DataTable schematable = base.connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new Object[] { null, null, null, "TABLE" });
            return schematable;
        }

        //public bool IsEqualSheetName(string sheetname)
        //{
        //    DataTable sheetNameList = GetExcelSheetName();

        //    for (int index = 0; index < sheetNameList.Rows.Count; index++)
        //    {
        //        if (sheetNameList.Rows[index]["TABLE_NAME"].ToString().ToUpper().Equals(sheetname.ToUpper()))
        //            return true;
        //    }

        //    return false;
        //}

        public string IsContainSheetName(string sheetname)
        {
            DataTable sheetNameList = GetExcelSheetName();

            for (int index = 0; index < sheetNameList.Rows.Count; index++)
            {
                if (sheetNameList.Rows[index]["TABLE_NAME"].ToString().ToUpper().Contains(sheetname.ToUpper()))
                    return sheetNameList.Rows[index]["TABLE_NAME"].ToString();
            }

            return string.Empty;
        }

        public DataSet ReadExcelFile(string sheetname)
        {
            string sqlString = string.Empty;

            if (sheetname.Contains("$"))
                sqlString = String.Format("SELECT * FROM [{0}];", sheetname);
            else
                sqlString = String.Format("SELECT * FROM [{0}$];", sheetname);

            return this.ExecuteQuery(sqlString);
        }

        public DataSet ReadExcelFile(string sheetname, string fieldname)
        {
            string sqlString = string.Empty;

            if (sheetname.Contains("$"))
                sqlString = String.Format("SELECT {0} FROM [{1}];", fieldname, sheetname);
            else
                sqlString = String.Format("SELECT {0} FROM [{1}$];", fieldname, sheetname);

            return this.ExecuteQuery(sqlString);
        }

        public DataTable ReadAgingReportFile(string sheetname)
        {
            return ReadExcelFile(sheetname, "FGA, Age, Pcs").Tables[0];
        }

        public DataTable ReadProfileReportFile(string sheetname)
        {
            return ReadExcelFile(sheetname).Tables[0];
        }

        public DataTable ReadOnhandReportFile(string sheetname)
        {
            return ReadExcelFile(sheetname, "FGAID, Region, On_Hand").Tables[0];
        }

        public DataTable ReadInTransitReportFile(string sheetname)
        {
            return ReadExcelFile(sheetname, "FGAID, Shipmode, Quantity, Region").Tables[0];
        }

        public DataTable ReadAPJSnPPartFilePartList(string sheetname)
        {
            return ReadExcelFile(sheetname, "SiteName, PartNumber, Category, UnitCost").Tables[0];
        }

        public DataTable ReadAPJSnPPartFileSiteNameList(string sheetname)
        {
            return ReadExcelFile(sheetname, "Masloc, SiteName").Tables[0];
        }

        public DataTable ReadIND3PLReportFile(string sheetname)
        {
            return ReadExcelFile(sheetname, "Owner, [Vendor Code], [Product Code], [Product Desc], [GRN Date], [Avail Qty]").Tables[0];
        }

        public DataTable ReadANZ3PLReportFile(string sheetname)
        {
            return ReadExcelFile(sheetname).Tables[0];
        }

        public DataTable ReadMST3PLReportFile(string sheetname)
        {
            return ReadExcelFile(sheetname, "Owner, Vendor, [Product Code], [Product Desc], [Age (Days)], [Phy Qty]").Tables[0];
        }

        public DataTable ReadJPN3PLReportFile(string sheetname)
        {
            return ReadExcelFile(sheetname).Tables[0];
        }

        public DataTable ReadCCCDragonReportFile(string sheetname)
        {
            return ReadExcelFile(sheetname, "PART_NUM, MAS_LOC, VENDOR_CODE, DMI_GOODPART").Tables[0];
        }
    }
}
