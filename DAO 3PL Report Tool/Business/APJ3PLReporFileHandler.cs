using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DAO_3PL_Report_Tool
{
    public class APJ3PLReporFileHandler
    {
        private UserSelectedValue UserSelectedValue = null;
        private DataTable ConsolidationAgingDataTable = null;
        private DataTable APJSnPPartListDataTable = null;
        private DataTable APJSnPSiteNameListDataTable = null;
        private Dictionary<string, string> ReportFileNames = null;
        private List<DataTable> APJ3PLReportDataTableList = new List<DataTable>();

        public APJ3PLReporFileHandler(UserSelectedValue userselectedvalue)
        {
            this.UserSelectedValue = userselectedvalue;
        }

        public void Process()
        {
            SearchReportFiles();
            LoadReportFiles();
            BuildDataTableStructure();
            FillRawData();            
            ExportConsolidationAgingReport();

            this.UserSelectedValue = null;
            this.ConsolidationAgingDataTable = null;
            this.APJSnPPartListDataTable = null;
            this.APJSnPSiteNameListDataTable = null;
            this.ReportFileNames = null;
            this.APJ3PLReportDataTableList = null;
            this.APJSnPSiteNameListDataTable = null;
        }

        #region ----- Search report files that need to be processed -----
        private void SearchReportFiles()
        {
            DirectoryInfo dir = null;
            string[] siteNameList = ConfigFileUtility.GetValue("SiteNameList").Split(',');

            ReportFileNames = new Dictionary<string, string>();
            foreach (string siteNme in siteNameList)
            {
                ReportFileNames.Add(siteNme, "");
            }

            string anz3PLReportPrefixName = ConfigFileUtility.GetValue("ANZ3PLReportPrefixName");
            string ind3PLReportPrefixName = ConfigFileUtility.GetValue("IND3PLReportPrefixName");
            string mst3PLReportPrefixName = ConfigFileUtility.GetValue("MST3PLReportPrefixName");
            string jpn3PLReportPrefixName = ConfigFileUtility.GetValue("JPN3PLReportPrefixName");
            string cccDragonReportPrefixName = ConfigFileUtility.GetValue("CCCDragonReportPrefixName");
            
            try
            {
                MiscUtility.LogHistory("Starting to search 3PL report files...");
                MainForm.backgroundworker.ReportProgress(0,
                    string.Format("[{0}] - Starting to search 3PL report files...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));

                dir = new DirectoryInfo(UserSelectedValue.APJ3PLReportSourceFolder);

                foreach (FileInfo fi in dir.GetFiles())
                {
                    if (fi.Name.Contains(anz3PLReportPrefixName))
                        ReportFileNames["ANZ"] = fi.FullName;

                    if (fi.Name.Contains(ind3PLReportPrefixName))
                        ReportFileNames["IND"] = fi.FullName;

                    if (fi.Name.Contains(mst3PLReportPrefixName))
                        ReportFileNames["MST"] = fi.FullName;

                    if (fi.Name.Contains(jpn3PLReportPrefixName))
                        ReportFileNames["JPN"] = fi.FullName;

                    if (fi.Name.Contains(cccDragonReportPrefixName))
                        ReportFileNames["CCC"] = fi.FullName;
                }

                foreach (KeyValuePair<string, string> item in ReportFileNames)
                {
                    if (item.Value.Length == 0)
                    {
                        MiscUtility.LogHistory(string.Format("Cannot find 3PL report file for {0}!", item.Key));
                        throw new Exception(string.Format("Cannot find 3PL report file for {0}!", item.Key));
                    }
                }

                MiscUtility.LogHistory(string.Format("Done!"));
                MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Function name: <SearchReportFiles>, Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }
        #endregion

        private void LoadReportFiles()
        {
            LoadAPJSnPPartListFile();

            LoadANZ3PLReport();
            LoadMST3PLReport();
            LoadIND3PLReport();
            LoadJapan3PLReport();
            LoadCCCDragonReport();            
        }

        #region ----- Load ANZ 3PL report -----
        private void LoadANZ3PLReport()
        {
            string fileName = ReportFileNames["ANZ"];
            string sheetName = ConfigFileUtility.GetValue("ANZ3PLReportSheetName");

            MiscUtility.LogHistory(string.Format("Starting to load ANZ 3PL file..."));
            MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Starting to load ANZ 3PL file...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));

            try
            {
                using (ExcelAccessDAO dao = new ExcelAccessDAO(fileName))
                {
                    string fullSheetName = dao.IsContainSheetName(sheetName);
                    if (fullSheetName.Contains(fullSheetName))
                    {
                        DataTable datatable = dao.ReadANZ3PLReportFile(fullSheetName);
                        datatable.TableName = "ANZ";
                        APJ3PLReportDataTableList.Add(datatable);
                    }
                    else
                    {
                        MiscUtility.LogHistory(string.Format("No found the sheet name - {0} in ANZ 3PL file {1}!", fullSheetName, fileName));
                        throw new Exception(string.Format("No found the sheet name - {0} in ANZ 3PL file {1}!", fullSheetName, fileName));
                    }
                }

                MiscUtility.LogHistory("Done!");
                MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("{0}, {1}", ex.Message, ex.StackTrace));
                throw;
            }
        }
        #endregion

        #region ----- Load MSN 3PL report -----
        private void LoadMST3PLReport()
        {
            string fileName = ReportFileNames["MST"];
            string sheetName = ConfigFileUtility.GetValue("MST3PLReportSheetName");

            MiscUtility.LogHistory(string.Format("Starting to load MST 3PL file..."));
            MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Starting to load MST 3PL file...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));

            try
            {
                using (ExcelAccessDAO dao = new ExcelAccessDAO(fileName))
                {
                    string fullSheetName = dao.IsContainSheetName(sheetName);
                    if (fullSheetName.Contains(fullSheetName))
                    {
                        DataTable datatable = dao.ReadMST3PLReportFile(fullSheetName);
                        datatable.TableName = "MST";
                        APJ3PLReportDataTableList.Add(datatable);
                    }
                    else
                    {
                        MiscUtility.LogHistory(string.Format("No found the sheet name - {0} in MST 3PL file {1}!", fullSheetName, fileName));
                        throw new Exception(string.Format("No found the sheet name - {0} in MST 3PL file {1}!", fullSheetName, fileName));
                    }
                }

                MiscUtility.LogHistory("Done!");
                MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("{0}, {1}", ex.Message, ex.StackTrace));
                throw;
            }
        }
        #endregion

        #region ----- Load India 3PL report -----
        private void LoadIND3PLReport()
        {
            string fileName = ReportFileNames["IND"];
            string sheetName = ConfigFileUtility.GetValue("IND3PLReportSheetName");

            MiscUtility.LogHistory(string.Format("Starting to load IND 3PL file..."));
            MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Starting to load IND 3PL file...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));

            try
            {
                using (ExcelAccessDAO dao = new ExcelAccessDAO(fileName))
                {
                    string fullSheetName = dao.IsContainSheetName(sheetName);
                    if (fullSheetName.Contains(fullSheetName))
                    {
                        DataTable datatable = dao.ReadIND3PLReportFile(fullSheetName);
                        datatable.TableName = "IND";
                        APJ3PLReportDataTableList.Add(datatable);
                    }
                    else
                    {
                        MiscUtility.LogHistory(string.Format("No found the sheet name - {0} in IND 3PL file {1}!", fullSheetName, fileName));
                        throw new Exception(string.Format("No found the sheet name - {0} in IND 3PL file {1}!", fullSheetName, fileName));
                    }
                }

                MiscUtility.LogHistory("Done!");
                MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("{0}, {1}", ex.Message, ex.StackTrace));
                throw;
            }
        }
        #endregion

        #region ----- Load CCC Dragon report -----
        private void LoadCCCDragonReport()
        {
            string fileName = ReportFileNames["CCC"];
            string sheetName = ConfigFileUtility.GetValue("CCCDragonReportSheetName");
            
            MiscUtility.LogHistory(string.Format("Starting to load CCC 3PL file..."));
            MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Starting to load CCC 3PL file...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));

            try
            {
                using (ExcelAccessDAO dao = new ExcelAccessDAO(fileName))
                {
                    string fullSheetName = dao.IsContainSheetName(sheetName);
                    if (fullSheetName.Contains(fullSheetName))
                    {
                        DataTable datatable = dao.ReadCCCDragonReportFile(fullSheetName);
                        datatable.TableName = "CCC";
                        APJ3PLReportDataTableList.Add(datatable);
                    }
                    else
                    {
                        MiscUtility.LogHistory(string.Format("No found the sheet name - {0} in CCC Dragon file {1}!", fullSheetName, fileName));
                        throw new Exception(string.Format("No found the sheet name - {0} in CCC Dragon file {1}!", fullSheetName, fileName));
                    }
                }

                MiscUtility.LogHistory("Done!");
                MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("{0}, {1}", ex.Message, ex.StackTrace));
                throw;
            }
        }
        #endregion

        #region ----- Load Japan 3PL report -----
        private void LoadJapan3PLReport()
        {
            string fileName = ReportFileNames["JPN"];
            string sheetName = ConfigFileUtility.GetValue("JPN3PLReportSheetName");

            MiscUtility.LogHistory(string.Format("Starting to load Japan 3PL file..."));
            MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Starting to load Japan 3PL file...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));

            try
            {
                using (ExcelAccessDAO dao = new ExcelAccessDAO(fileName))
                {
                    string fullSheetName = dao.IsContainSheetName(sheetName);
                    if (fullSheetName.Contains(fullSheetName))
                    {
                        DataTable datatable = dao.ReadJPN3PLReportFile(fullSheetName);
                        datatable.TableName = "JPN";
                        APJ3PLReportDataTableList.Add(datatable);
                    }
                    else
                    {
                        MiscUtility.LogHistory(string.Format("No found the sheet name - {0} in Japan 3PL file {1}!", fullSheetName, fileName));
                        throw new Exception(string.Format("No found the sheet name - {0} in Japan 3PL file {1}!", fullSheetName, fileName));
                    }
                }

                MiscUtility.LogHistory("Done!");
                MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("{0}, {1}", ex.Message, ex.StackTrace));
                throw;
            }
        }
        #endregion

        #region ----- Load SnP Part List file -----
        private void LoadAPJSnPPartListFile()
        {
            string apjSnpPartListFile = ConfigFileUtility.GetValue("APJSnPPartListFile");
            string partListSheetName = ConfigFileUtility.GetValue("APJSnPPartListFilePartListSheetName");
            string siteNameListSheetName = ConfigFileUtility.GetValue("APJSnPPartListFileSiteNameListSheetName");

            if (!File.Exists(apjSnpPartListFile))
            {
                MiscUtility.LogHistory(string.Format("Cannot find APJ SnP Part List file - {0}!", apjSnpPartListFile));
                throw new Exception(string.Format("Cannot find APJ SnP Part List file - {0}!", apjSnpPartListFile));
            }

            MiscUtility.LogHistory(string.Format("Starting to search APJ SnP Part List file..."));
            MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Starting to search SnP Part List file...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));

            try
            {
                using (ExcelAccessDAO dao = new ExcelAccessDAO(apjSnpPartListFile))
                {
                    string fullSheetName = dao.IsContainSheetName(partListSheetName);
                    if (fullSheetName.Contains(fullSheetName))
                    {
                        APJSnPPartListDataTable = dao.ReadAPJSnPPartFilePartList(fullSheetName);
                    }
                    else
                    {
                        MiscUtility.LogHistory(string.Format("No found the sheet name - {0} from APJ SnP Part List file {1}!", fullSheetName, apjSnpPartListFile));
                        throw new Exception(string.Format("No found the sheet name - {0} from APJ SnP Part List file {1}!", fullSheetName, apjSnpPartListFile));
                    }

                    fullSheetName = dao.IsContainSheetName(siteNameListSheetName);
                    if (fullSheetName.Contains(fullSheetName))
                    {
                        APJSnPSiteNameListDataTable = dao.ReadAPJSnPPartFileSiteNameList(fullSheetName);
                    }
                    else
                    {
                        MiscUtility.LogHistory(string.Format("No found the sheet name - {0} APJ from SnP Part List file {1}!", fullSheetName, apjSnpPartListFile));
                        throw new Exception(string.Format("No found the sheet name - {0} APJ from SnP Part List file {1}!", fullSheetName, apjSnpPartListFile));
                    }
                }

                MiscUtility.LogHistory("Done!");
                MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("{0}, {1}", ex.Message, ex.StackTrace));
                throw;
            }
        }
        #endregion

        #region ----- Build a struction of data table -----
        private void BuildDataTableStructure()
        {
            ConsolidationAgingDataTable = new DataTable();

            DataColumn dc = new DataColumn();
            dc.ColumnName = "Date";
            dc.DataType = System.Type.GetType("System.String");
            ConsolidationAgingDataTable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Mas Loc";
            dc.DataType = System.Type.GetType("System.String");
            ConsolidationAgingDataTable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Vendor Code";
            dc.DataType = System.Type.GetType("System.String");
            ConsolidationAgingDataTable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Site Name";
            dc.DataType = System.Type.GetType("System.String");
            ConsolidationAgingDataTable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Part Number";
            dc.DataType = System.Type.GetType("System.String");
            ConsolidationAgingDataTable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Description";
            dc.DataType = System.Type.GetType("System.String");
            ConsolidationAgingDataTable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Category";
            dc.DataType = System.Type.GetType("System.String");
            ConsolidationAgingDataTable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Qty";
            dc.DataType = System.Type.GetType("System.Int32");
            ConsolidationAgingDataTable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Unit Cost";
            dc.DataType = System.Type.GetType("System.Double");
            ConsolidationAgingDataTable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Total Cost";
            dc.DataType = System.Type.GetType("System.Double");
            ConsolidationAgingDataTable.Columns.Add(dc);

            /*
            dc = new DataColumn();
            dc.ColumnName = "<= 15 Days";
            dc.DataType = System.Type.GetType("System.Int32");
            ConsolidationAgingDataTable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = ">15 and <=30 Days";
            dc.DataType = System.Type.GetType("System.Int32");
            ConsolidationAgingDataTable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = ">30 and <=45 Days";
            dc.DataType = System.Type.GetType("System.Int32");
            ConsolidationAgingDataTable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = ">45 and <=60 Days";
            dc.DataType = System.Type.GetType("System.Int32");
            ConsolidationAgingDataTable.Columns.Add(dc);
            */

            dc = new DataColumn();
            dc.ColumnName = "<= 30 Days";
            dc.DataType = System.Type.GetType("System.Int32");
            ConsolidationAgingDataTable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = ">30 and <=60 Days";
            dc.DataType = System.Type.GetType("System.Int32");
            ConsolidationAgingDataTable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = ">60 and <=90 Days";
            dc.DataType = System.Type.GetType("System.Int32");
            ConsolidationAgingDataTable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = ">90 and <=120 Days";
            dc.DataType = System.Type.GetType("System.Int32");
            ConsolidationAgingDataTable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = ">120 and <=150 Days";
            dc.DataType = System.Type.GetType("System.Int32");
            ConsolidationAgingDataTable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = ">150 and <=180 Days";
            dc.DataType = System.Type.GetType("System.Int32");
            ConsolidationAgingDataTable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = ">180 and <=210 Days";
            dc.DataType = System.Type.GetType("System.Int32");
            ConsolidationAgingDataTable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = ">210 Days";
            dc.DataType = System.Type.GetType("System.Int32");
            ConsolidationAgingDataTable.Columns.Add(dc);
        }
        #endregion

        private void FillRawData()
        {
            FillAgingData();
            FillCategoryAndSiteNameData();
        }

        #region ----- Fill 3PL aging raw data into ConsolidationAgingDataTable -----
        private void FillAgingData()
        {
            bool flag = false;
            DataRow dataRow = null;            
            string masLoc = string.Empty;
            string partNumber = string.Empty;
            string siteCode = string.Empty;
            string currentDate = DateTime.Now.Date.ToShortDateString();

            for (int num = 0; num < this.APJ3PLReportDataTableList.Count; num++)
            {
                #region ----- Fill 3PL ANZ inventory again raw data -----
                if (APJ3PLReportDataTableList[num].TableName.Equals("ANZ"))
                {
                    MiscUtility.LogHistory(string.Format("Starting to fill 3PL ANZ raw data..."));
                    MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Starting to fill 3PL ANZ raw data...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));

                    for (int index = 0; index < this.APJ3PLReportDataTableList[num].Rows.Count; index++)
                    {
                        flag = false;
                        siteCode = this.APJ3PLReportDataTableList[num].Rows[index]["SiteCode  "].ToString().Trim();
                        partNumber = this.APJ3PLReportDataTableList[num].Rows[index]["Sku                 "].ToString().Trim();
                        masLoc = this.APJ3PLReportDataTableList[num].Rows[index]["facil"].ToString().Trim();

                        // Filter out non-DMI part from ANZ 3PL part list
                        //if (siteCode.Equals("DMI") && partNumber.Length == 5)
                        if (partNumber.Length == 5)
                        {
                            for (int indey = 0; indey < this.ConsolidationAgingDataTable.Rows.Count; indey++)
                            {
                                if (masLoc.Equals(this.ConsolidationAgingDataTable.Rows[indey]["Mas Loc"])
                                    && partNumber.Equals(this.ConsolidationAgingDataTable.Rows[indey]["Part Number"]))
                                {
                                    /*
                                    this.ConsolidationAgingDataTable.Rows[indey]["<= 15 Days"] =
                                        Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["<= 15 Days"]) 
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["<= 15 Days"]);

                                    this.ConsolidationAgingDataTable.Rows[indey][">15 and <=30 Days"] =
                                        Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">15 and <=30 Days"]) 
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index][">15 and <=30 Days   "]);

                                    this.ConsolidationAgingDataTable.Rows[indey][">30 and <=45 Days"] =
                                        Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">30 and <=45 Days"]) 
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index][">30 and <=45 Days   "]);

                                    this.ConsolidationAgingDataTable.Rows[indey][">45 and <=60 Days"] =
                                        Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">45 and <=60 Days"]) 
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index][">45 and <=60 Days   "]);
                                    */

                                    this.ConsolidationAgingDataTable.Rows[indey]["<= 30 Days"] =
                                        Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["<= 30 Days"])
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["<= 15 Days"])
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index][">15 and <=30 Days   "]);

                                    this.ConsolidationAgingDataTable.Rows[indey][">30 and <=60 Days"] =
                                        Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">30 and <=60 Days"])
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index][">30 and <=45 Days   "])
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index][">45 and <=60 Days   "]);

                                    this.ConsolidationAgingDataTable.Rows[indey][">60 and <=90 Days"] =
                                        Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">60 and <=90 Days"]) 
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index][">60 and <=90 Days   "]);

                                    this.ConsolidationAgingDataTable.Rows[indey][">90 and <=120 Days"] =
                                        Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">90 and <=120 Days"]) 
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index][">90 and <=120 Days  "]);

                                    this.ConsolidationAgingDataTable.Rows[indey][">120 and <=150 Days"] =
                                        Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">120 and <=150 Days"]) 
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index][">120 and <=150 Days "]);

                                    this.ConsolidationAgingDataTable.Rows[indey][">150 and <=180 Days"] =
                                        Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">150 and <=180 Days"]) 
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index][">150 and <=180 Days "]);

                                    this.ConsolidationAgingDataTable.Rows[indey][">180 and <=210 Days"] =
                                        Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">180 and <=210 Days"]) 
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index][">180 and <=210 Days "]);

                                    this.ConsolidationAgingDataTable.Rows[indey][">210 Days"] =
                                        Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">210 Days"]) 
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index][">210 Days "]);

                                    this.ConsolidationAgingDataTable.Rows[indey]["Qty"] = 
                                        /*Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["<= 15 Days"])
                                        + Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">15 and <=30 Days"])
                                        + Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">30 and <=45 Days"])
                                        + Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">45 and <=60 Days"])*/

                                        Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["<= 30 Days"])
                                        + Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">30 and <=60 Days"])
                                        + Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">60 and <=90 Days"])
                                        + Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">90 and <=120 Days"])
                                        + Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">120 and <=150 Days"])
                                        + Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">150 and <=180 Days"])
                                        + Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">180 and <=210 Days"])
                                        + Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">210 Days"]);

                                    flag = true;
                                    break;
                                }
                            }

                            if (!flag)
                            {
                                dataRow = this.ConsolidationAgingDataTable.NewRow();

                                dataRow["Date"] = currentDate;
                                dataRow["Mas Loc"] = masLoc;
                                dataRow["Vendor Code"] = this.APJ3PLReportDataTableList[num].Rows[index]["Storer         "].ToString().Trim();
                                dataRow["Part Number"] = partNumber;
                                dataRow["Category"] = "";                                
                                /*dataRow["<= 15 Days"] = this.APJ3PLReportDataTableList[num].Rows[index]["<= 15 Days"].ToString().Trim();
                                dataRow[">15 and <=30 Days"] = this.APJ3PLReportDataTableList[num].Rows[index][">15 and <=30 Days   "].ToString().Trim();
                                dataRow[">30 and <=45 Days"] = this.APJ3PLReportDataTableList[num].Rows[index][">30 and <=45 Days   "].ToString().Trim();
                                dataRow[">45 and <=60 Days"] = this.APJ3PLReportDataTableList[num].Rows[index][">45 and <=60 Days   "].ToString().Trim();*/

                                dataRow["<= 30 Days"] = 
                                    Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["<= 15 Days"])
                                    + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index][">15 and <=30 Days   "]);

                                dataRow[">30 and <=60 Days"] =
                                    Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index][">30 and <=45 Days   "])
                                    + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index][">45 and <=60 Days   "]);

                                dataRow[">60 and <=90 Days"] = this.APJ3PLReportDataTableList[num].Rows[index][">60 and <=90 Days   "].ToString().Trim();
                                dataRow[">90 and <=120 Days"] = this.APJ3PLReportDataTableList[num].Rows[index][">90 and <=120 Days  "].ToString().Trim();
                                dataRow[">120 and <=150 Days"] = this.APJ3PLReportDataTableList[num].Rows[index][">120 and <=150 Days "].ToString().Trim();
                                dataRow[">150 and <=180 Days"] = this.APJ3PLReportDataTableList[num].Rows[index][">150 and <=180 Days "].ToString().Trim();
                                dataRow[">180 and <=210 Days"] = this.APJ3PLReportDataTableList[num].Rows[index][">180 and <=210 Days "].ToString().Trim();
                                dataRow[">210 Days"] = this.APJ3PLReportDataTableList[num].Rows[index][">210 Days "].ToString().Trim();
                                dataRow["Qty"] = 
                                    /*Convert.ToInt32(dataRow["<= 15 Days"])
                                    + Convert.ToInt32(dataRow[">15 and <=30 Days"])
                                    + Convert.ToInt32(dataRow[">30 and <=45 Days"])
                                    + Convert.ToInt32(dataRow[">45 and <=60 Days"])*/
                                    
                                    Convert.ToInt32(dataRow["<= 30 Days"])
                                    + Convert.ToInt32(dataRow[">30 and <=60 Days"])
                                    + Convert.ToInt32(dataRow[">60 and <=90 Days"])
                                    + Convert.ToInt32(dataRow[">90 and <=120 Days"])
                                    + Convert.ToInt32(dataRow[">120 and <=150 Days"])
                                    + Convert.ToInt32(dataRow[">150 and <=180 Days"])
                                    + Convert.ToInt32(dataRow[">180 and <=210 Days"])
                                    + Convert.ToInt32(dataRow[">210 Days"]);

                                this.ConsolidationAgingDataTable.Rows.Add(dataRow);
                            }
                        }
                    }
                    
                    MiscUtility.LogHistory("Done!");
                    MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
                }                                
                #endregion

                #region ----- Fill 3PL MST inventory again raw data -----                
                int ageDays = 0;
                int quantity = 0;

                if (APJ3PLReportDataTableList[num].TableName.Equals("MST"))
                {
                    MiscUtility.LogHistory(string.Format("Starting to fill 3PL MST raw data..."));
                    MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Starting to fill 3PL MST raw data...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));

                    for (int index = 0; index < this.APJ3PLReportDataTableList[num].Rows.Count; index++)
                    {
                        flag = false;
                        
                        masLoc = this.APJ3PLReportDataTableList[num].Rows[index]["Owner"].ToString().ToUpper().Trim();
                        partNumber = this.APJ3PLReportDataTableList[num].Rows[index]["Product Code"].ToString().Trim();
                        ageDays = Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["Age (Days)"]);
                        quantity = Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["Phy Qty"]);

                        // Filter out non-DMI part from MST 3PL part list
                        //if (masLoc.Equals("DELLMST"))
                        //{
                        if (masLoc.Equals("DELLGLS")) continue;

                            for (int indey = 0; indey < this.ConsolidationAgingDataTable.Rows.Count; indey++)
                            {
                                if (masLoc.Equals(this.ConsolidationAgingDataTable.Rows[indey]["Mas Loc"])
                                        && partNumber.Equals(this.ConsolidationAgingDataTable.Rows[indey]["Part Number"]))
                                {
                                    /*
                                    if (ageDays <= 15)
                                    {
                                        this.ConsolidationAgingDataTable.Rows[indey]["<= 15 Days"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["<= 15 Days"]) + quantity;
                                        this.ConsolidationAgingDataTable.Rows[indey]["Qty"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["Qty"]) + quantity;
                                    }
                                    else if (ageDays > 15 && ageDays <= 30)
                                    {
                                        this.ConsolidationAgingDataTable.Rows[indey][">15 and <=30 Days"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">15 and <=30 Days"]) + quantity;
                                        this.ConsolidationAgingDataTable.Rows[indey]["Qty"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["Qty"]) + quantity;
                                    }
                                    else if (ageDays > 30 && ageDays <= 45)
                                    {
                                        this.ConsolidationAgingDataTable.Rows[indey][">30 and <=45 Days"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">30 and <=45 Days"]) + quantity;
                                        this.ConsolidationAgingDataTable.Rows[indey]["Qty"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["Qty"]) + quantity;
                                    }
                                    else if (ageDays > 45 && ageDays <= 60)
                                    {
                                        this.ConsolidationAgingDataTable.Rows[indey][">45 and <=60 Days"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">45 and <=60 Days"]) + quantity;
                                        this.ConsolidationAgingDataTable.Rows[indey]["Qty"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["Qty"]) + quantity;
                                    }
                                     */

                                    if (ageDays <= 30)
                                    {
                                        this.ConsolidationAgingDataTable.Rows[indey]["<= 30 Days"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["<= 30 Days"]) + quantity;
                                        this.ConsolidationAgingDataTable.Rows[indey]["Qty"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["Qty"]) + quantity;
                                    }
                                    else if (ageDays > 30 && ageDays <= 60)
                                    {
                                        this.ConsolidationAgingDataTable.Rows[indey][">30 and <=60 Days"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">30 and <=60 Days"]) + quantity;
                                        this.ConsolidationAgingDataTable.Rows[indey]["Qty"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["Qty"]) + quantity;
                                    }
                                    else if (ageDays > 60 && ageDays <= 90)
                                    {
                                        this.ConsolidationAgingDataTable.Rows[indey][">60 and <=90 Days"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">60 and <=90 Days"]) + quantity;
                                        this.ConsolidationAgingDataTable.Rows[indey]["Qty"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["Qty"]) + quantity;
                                    }
                                    else if (ageDays > 90 && ageDays <= 120)
                                    {
                                        this.ConsolidationAgingDataTable.Rows[indey][">90 and <=120 Days"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">90 and <=120 Days"]) + quantity;
                                        this.ConsolidationAgingDataTable.Rows[indey]["Qty"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["Qty"]) + quantity;
                                    }
                                    else if (ageDays > 120 && ageDays <= 150)
                                    {
                                        this.ConsolidationAgingDataTable.Rows[indey][">120 and <=150 Days"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">120 and <=150 Days"]) + quantity;
                                        this.ConsolidationAgingDataTable.Rows[indey]["Qty"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["Qty"]) + quantity;
                                    }
                                    else if (ageDays > 150 && ageDays <= 180)
                                    {
                                        this.ConsolidationAgingDataTable.Rows[indey][">150 and <=180 Days"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">150 and <=180 Days"]) + quantity;
                                        this.ConsolidationAgingDataTable.Rows[indey]["Qty"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["Qty"]) + quantity;
                                    }
                                    else if (ageDays > 180 && ageDays <= 120)
                                    {
                                        this.ConsolidationAgingDataTable.Rows[indey][">180 and <=210 Days"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">180 and <=210 Days"]) + quantity;
                                        this.ConsolidationAgingDataTable.Rows[indey]["Qty"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["Qty"]) + quantity;
                                    }
                                    else if (ageDays > 210)
                                    {
                                        this.ConsolidationAgingDataTable.Rows[indey][">210 Days"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">210 Days"]) + quantity;
                                        this.ConsolidationAgingDataTable.Rows[indey]["Qty"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["Qty"]) + quantity;
                                    }

                                    flag = true;
                                    break;
                                }
                            }

                            if (!flag)
                            {
                                dataRow = this.ConsolidationAgingDataTable.NewRow();

                                dataRow["Date"] = currentDate;
                                dataRow["Mas Loc"] = masLoc;
                                dataRow["Vendor Code"] = this.APJ3PLReportDataTableList[num].Rows[index]["Vendor"].ToString().Trim();
                                dataRow["Part Number"] = partNumber;
                                dataRow["Category"] = "";
                                /*
                                dataRow["<= 15 Days"] = 0;
                                dataRow[">15 and <=30 Days"] = 0;
                                dataRow[">30 and <=45 Days"] = 0;
                                dataRow[">45 and <=60 Days"] = 0;
                                 * */
                                dataRow["<= 30 Days"] = 0;
                                dataRow[">30 and <=60 Days"] = 0;
                                dataRow[">60 and <=90 Days"] = 0;
                                dataRow[">90 and <=120 Days"] = 0;
                                dataRow[">120 and <=150 Days"] = 0;
                                dataRow[">150 and <=180 Days"] = 0;
                                dataRow[">180 and <=210 Days"] = 0;
                                dataRow[">210 Days"] = 0;

                                /*
                                if (ageDays <= 15)
                                    dataRow["<= 15 Days"] = quantity;

                                else if (ageDays > 15 && ageDays <= 30)
                                    dataRow[">15 and <=30 Days"] = quantity;

                                else if (ageDays > 30 && ageDays <= 45)
                                    dataRow[">30 and <=45 Days"] = quantity;

                                else if (ageDays > 45 && ageDays <= 60)
                                    dataRow[">45 and <=60 Days"] = quantity;
                                */

                                if (ageDays <= 30)
                                    dataRow["<= 30 Days"] = quantity;

                                else if (ageDays > 30 && ageDays <= 60)
                                    dataRow[">30 and <=60 Days"] = quantity;

                                else if (ageDays > 60 && ageDays <= 90)
                                    dataRow[">60 and <=90 Days"] = quantity;

                                else if (ageDays > 90 && ageDays <= 120)
                                    dataRow[">90 and <=120 Days"] = quantity;

                                else if (ageDays > 120 && ageDays <= 150)
                                    dataRow[">120 and <=150 Days"] = quantity;

                                else if (ageDays > 150 && ageDays <= 180)
                                    dataRow[">150 and <=180 Days"] = quantity;

                                else if (ageDays > 180 && ageDays <= 120)
                                    dataRow[">180 and <=210 Days"] = quantity;

                                else if (ageDays > 210)
                                    dataRow[">210 Days"] = quantity;

                                dataRow["Qty"] = 
                                    /*Convert.ToInt32(dataRow["<= 15 Days"])
                                    + Convert.ToInt32(dataRow[">15 and <=30 Days"])
                                    + Convert.ToInt32(dataRow[">30 and <=45 Days"])
                                    + Convert.ToInt32(dataRow[">45 and <=60 Days"])
                                     */
                                    Convert.ToInt32(dataRow["<= 30 Days"])
                                    + Convert.ToInt32(dataRow[">30 and <=60 Days"])
                                    + Convert.ToInt32(dataRow[">60 and <=90 Days"])
                                    + Convert.ToInt32(dataRow[">90 and <=120 Days"])
                                    + Convert.ToInt32(dataRow[">120 and <=150 Days"])
                                    + Convert.ToInt32(dataRow[">150 and <=180 Days"])
                                    + Convert.ToInt32(dataRow[">180 and <=210 Days"])
                                    + Convert.ToInt32(dataRow[">210 Days"]);

                                this.ConsolidationAgingDataTable.Rows.Add(dataRow);
                            }
                        //}
                    }

                    MiscUtility.LogHistory("Done!");
                    MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
                }
                #endregion

                #region ----- Fill 3PL IND inventory again raw data -----
                string grnDate = string.Empty;
                string vendorCode = string.Empty;

                if (APJ3PLReportDataTableList[num].TableName.Equals("IND"))
                {
                    MiscUtility.LogHistory(string.Format("Starting to fill 3PL IND raw data..."));
                    MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Starting to fill 3PL IND raw data...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));

                    for (int index = 0; index < this.APJ3PLReportDataTableList[num].Rows.Count; index++)
                    {
                        flag = false;

                        vendorCode = this.APJ3PLReportDataTableList[num].Rows[index]["Vendor Code"].ToString().Trim().ToUpper();
                        masLoc = this.APJ3PLReportDataTableList[num].Rows[index]["Owner"].ToString().Trim();
                        partNumber = this.APJ3PLReportDataTableList[num].Rows[index]["Product Code"].ToString().Trim();
                        quantity = Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["Avail Qty"]);

                        grnDate = this.APJ3PLReportDataTableList[num].Rows[index]["GRN Date"].ToString().Substring(0, 10);
                        ageDays = (Int32)DateTime.Parse(currentDate).Subtract(DateTime.Parse(grnDate)).TotalDays;


                        //if (vendorCode.Equals("DELL"))
                        //{
                            for (int indey = 0; indey < this.ConsolidationAgingDataTable.Rows.Count; indey++)
                            {
                                if (masLoc.Equals(this.ConsolidationAgingDataTable.Rows[indey]["Mas Loc"])
                                        && partNumber.Equals(this.ConsolidationAgingDataTable.Rows[indey]["Part Number"]))
                                {
                                    /*
                                    if (ageDays <= 15)
                                    {
                                        this.ConsolidationAgingDataTable.Rows[indey]["<= 15 Days"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["<= 15 Days"]) + quantity;
                                        this.ConsolidationAgingDataTable.Rows[indey]["Qty"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["Qty"]) + quantity;
                                    }
                                    else if (ageDays > 15 && ageDays <= 30)
                                    {
                                        this.ConsolidationAgingDataTable.Rows[indey][">15 and <=30 Days"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">15 and <=30 Days"]) + quantity;
                                        this.ConsolidationAgingDataTable.Rows[indey]["Qty"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["Qty"]) + quantity;
                                    }
                                    else if (ageDays > 30 && ageDays <= 45)
                                    {
                                        this.ConsolidationAgingDataTable.Rows[indey][">30 and <=45 Days"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">30 and <=45 Days"]) + quantity;
                                        this.ConsolidationAgingDataTable.Rows[indey]["Qty"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["Qty"]) + quantity;
                                    }
                                    else if (ageDays > 45 && ageDays <= 60)
                                    {
                                        this.ConsolidationAgingDataTable.Rows[indey][">45 and <=60 Days"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">45 and <=60 Days"]) + quantity;
                                        this.ConsolidationAgingDataTable.Rows[indey]["Qty"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["Qty"]) + quantity;
                                    }
                                     */

                                    if (ageDays <= 30)
                                    {
                                        this.ConsolidationAgingDataTable.Rows[indey]["<= 30 Days"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["<= 30 Days"]) + quantity;
                                        this.ConsolidationAgingDataTable.Rows[indey]["Qty"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["Qty"]) + quantity;
                                    }
                                    else if (ageDays > 30 && ageDays <= 60)
                                    {
                                        this.ConsolidationAgingDataTable.Rows[indey][">30 and <=60 Days"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">30 and <=60 Days"]) + quantity;
                                        this.ConsolidationAgingDataTable.Rows[indey]["Qty"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["Qty"]) + quantity;
                                    }
                                    else if (ageDays > 60 && ageDays <= 90)
                                    {
                                        this.ConsolidationAgingDataTable.Rows[indey][">60 and <=90 Days"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">60 and <=90 Days"]) + quantity;
                                        this.ConsolidationAgingDataTable.Rows[indey]["Qty"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["Qty"]) + quantity;
                                    }
                                    else if (ageDays > 90 && ageDays <= 120)
                                    {
                                        this.ConsolidationAgingDataTable.Rows[indey][">90 and <=120 Days"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">90 and <=120 Days"]) + quantity;
                                        this.ConsolidationAgingDataTable.Rows[indey]["Qty"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["Qty"]) + quantity;
                                    }
                                    else if (ageDays > 120 && ageDays <= 150)
                                    {
                                        this.ConsolidationAgingDataTable.Rows[indey][">120 and <=150 Days"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">120 and <=150 Days"]) + quantity;
                                        this.ConsolidationAgingDataTable.Rows[indey]["Qty"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["Qty"]) + quantity;
                                    }
                                    else if (ageDays > 150 && ageDays <= 180)
                                    {
                                        this.ConsolidationAgingDataTable.Rows[indey][">150 and <=180 Days"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">150 and <=180 Days"]) + quantity;
                                        this.ConsolidationAgingDataTable.Rows[indey]["Qty"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["Qty"]) + quantity;
                                    }
                                    else if (ageDays > 180 && ageDays <= 120)
                                    {
                                        this.ConsolidationAgingDataTable.Rows[indey][">180 and <=210 Days"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">180 and <=210 Days"]) + quantity;
                                        this.ConsolidationAgingDataTable.Rows[indey]["Qty"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["Qty"]) + quantity;
                                    }
                                    else if (ageDays > 210)
                                    {
                                        this.ConsolidationAgingDataTable.Rows[indey][">210 Days"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">210 Days"]) + quantity;
                                        this.ConsolidationAgingDataTable.Rows[indey]["Qty"] =
                                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["Qty"]) + quantity;
                                    }

                                    flag = true;
                                    break;
                                }
                            }

                            if (!flag)
                            {
                                dataRow = this.ConsolidationAgingDataTable.NewRow();

                                dataRow["Date"] = currentDate;
                                dataRow["Mas Loc"] = masLoc;
                                dataRow["Vendor Code"] = vendorCode;
                                dataRow["Part Number"] = partNumber;
                                dataRow["Category"] = "";
                                /*dataRow["<= 15 Days"] = 0;
                                dataRow[">15 and <=30 Days"] = 0;
                                dataRow[">30 and <=45 Days"] = 0;
                                dataRow[">45 and <=60 Days"] = 0;*/

                                dataRow["<= 30 Days"] = 0;
                                dataRow[">30 and <=60 Days"] = 0;
                                dataRow[">60 and <=90 Days"] = 0;
                                dataRow[">90 and <=120 Days"] = 0;
                                dataRow[">120 and <=150 Days"] = 0;
                                dataRow[">150 and <=180 Days"] = 0;
                                dataRow[">180 and <=210 Days"] = 0;
                                dataRow[">210 Days"] = 0;

                                /*
                                if (ageDays <= 15)
                                    dataRow["<= 15 Days"] = quantity;

                                else if (ageDays > 15 && ageDays <= 30)
                                    dataRow[">15 and <=30 Days"] = quantity;

                                else if (ageDays > 30 && ageDays <= 45)
                                    dataRow[">30 and <=45 Days"] = quantity;

                                else if (ageDays > 45 && ageDays <= 60)
                                    dataRow[">45 and <=60 Days"] = quantity;
                                */

                                if (ageDays <= 30)
                                    dataRow["<= 30 Days"] = quantity;

                                else if (ageDays > 30 && ageDays <= 60)
                                    dataRow[">30 and <=60 Days"] = quantity;

                                else if (ageDays > 60 && ageDays <= 90)
                                    dataRow[">60 and <=90 Days"] = quantity;

                                else if (ageDays > 90 && ageDays <= 120)
                                    dataRow[">90 and <=120 Days"] = quantity;

                                else if (ageDays > 120 && ageDays <= 150)
                                    dataRow[">120 and <=150 Days"] = quantity;

                                else if (ageDays > 150 && ageDays <= 180)
                                    dataRow[">150 and <=180 Days"] = quantity;

                                else if (ageDays > 180 && ageDays <= 120)
                                    dataRow[">180 and <=210 Days"] = quantity;

                                else if (ageDays > 210)
                                    dataRow[">210 Days"] = quantity;

                                dataRow["Qty"] = 
                                    /*Convert.ToInt32(dataRow["<= 15 Days"])
                                    + Convert.ToInt32(dataRow[">15 and <=30 Days"])
                                    + Convert.ToInt32(dataRow[">30 and <=45 Days"])
                                    + Convert.ToInt32(dataRow[">45 and <=60 Days"])*/

                                    Convert.ToInt32(dataRow["<= 30 Days"])
                                    + Convert.ToInt32(dataRow[">30 and <=60 Days"])
                                    + Convert.ToInt32(dataRow[">60 and <=90 Days"])
                                    + Convert.ToInt32(dataRow[">90 and <=120 Days"])
                                    + Convert.ToInt32(dataRow[">120 and <=150 Days"])
                                    + Convert.ToInt32(dataRow[">150 and <=180 Days"])
                                    + Convert.ToInt32(dataRow[">180 and <=210 Days"])
                                    + Convert.ToInt32(dataRow[">210 Days"]);

                                this.ConsolidationAgingDataTable.Rows.Add(dataRow);
                            }
                        //}
                    }

                    MiscUtility.LogHistory("Done!");
                    MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
                }
                #endregion

                #region ----- Fill Japan 3PL inventory again raw data -----
                masLoc = "NRT-FDX";

                if (APJ3PLReportDataTableList[num].TableName.Equals("JPN"))
                {
                    MiscUtility.LogHistory(string.Format("Starting to fill 3PL JPN raw data..."));
                    MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Starting to fill 3PL JPN raw data...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));

                    for (int index = 0; index < this.APJ3PLReportDataTableList[num].Rows.Count; index++)
                    {
                        flag = false;
                        partNumber = this.APJ3PLReportDataTableList[num].Rows[index]["Part"].ToString().Trim();

                        if (partNumber.Length == 5)
                        {
                            for (int indey = 0; indey < this.ConsolidationAgingDataTable.Rows.Count; indey++)
                            {
                                if (masLoc.Equals(this.ConsolidationAgingDataTable.Rows[indey]["Mas Loc"])
                                        && partNumber.Equals(this.ConsolidationAgingDataTable.Rows[indey]["Part Number"]))
                                {
                                    this.ConsolidationAgingDataTable.Rows[indey]["<= 30 Days"] =
                                        Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey]["<= 30 days"])
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["<30 days"]);

                                    this.ConsolidationAgingDataTable.Rows[indey][">30 and <=60 Days"] =
                                        Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">30 and <=60 Days"])
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["30-60"]);

                                    this.ConsolidationAgingDataTable.Rows[indey][">60 and <=90 Days"] =
                                        Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">60 and <=90 Days"])
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["60-90"]);

                                    this.ConsolidationAgingDataTable.Rows[indey][">90 and <=120 Days"] =
                                        Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">90 and <=120 Days"])
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["90-120"]);

                                    this.ConsolidationAgingDataTable.Rows[indey][">120 and <=150 Days"] =
                                        Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">120 and <=150 Days"])
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["120-150"]);

                                    this.ConsolidationAgingDataTable.Rows[indey][">150 and <=180 Days"] =
                                        Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">150 and <=180 Days"])
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["150-180"]);

                                    this.ConsolidationAgingDataTable.Rows[indey][">180 and <=210 Days"] =
                                        Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">180 and <=210 Days"])
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["180-210"]);

                                    this.ConsolidationAgingDataTable.Rows[indey][">210 Days"] =
                                        Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[indey][">210 Days"])
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["210-240"])
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["240-270"])
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["270-300"])
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["300-330"])
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["330-360"])
                                        + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["Over 360"]);
                                }

                                if (!flag)
                                {
                                    dataRow = this.ConsolidationAgingDataTable.NewRow();

                                    dataRow["Date"] = currentDate;
                                    dataRow["Mas Loc"] = masLoc;
                                    dataRow["Vendor Code"] = "";
                                    dataRow["Part Number"] = partNumber;
                                    dataRow["Category"] = "";
                                    /*dataRow["<= 15 Days"] = 0;
                                    dataRow[">15 and <=30 Days"] = 0;
                                    dataRow[">30 and <=45 Days"] = 0;
                                    dataRow[">45 and <=60 Days"] = 0;*/

                                    dataRow["<= 30 Days"] = 0;
                                    dataRow[">30 and <=60 Days"] = 0;
                                    dataRow[">60 and <=90 Days"] = 0;
                                    dataRow[">90 and <=120 Days"] = 0;
                                    dataRow[">120 and <=150 Days"] = 0;
                                    dataRow[">150 and <=180 Days"] = 0;
                                    dataRow[">180 and <=210 Days"] = 0;
                                    dataRow[">210 Days"] = 0;

                                    /*
                                    if (ageDays <= 15)
                                        dataRow["<= 15 Days"] = quantity;

                                    else if (ageDays > 15 && ageDays <= 30)
                                        dataRow[">15 and <=30 Days"] = quantity;

                                    else if (ageDays > 30 && ageDays <= 45)
                                        dataRow[">30 and <=45 Days"] = quantity;

                                    else if (ageDays > 45 && ageDays <= 60)
                                        dataRow[">45 and <=60 Days"] = quantity;
                                    */

                                    if (this.APJ3PLReportDataTableList[num].Rows[index]["<30 days"].ToString().Trim().Length != 0)
                                        dataRow["<= 30 Days"] = Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["<30 days"]);
                                    else
                                        dataRow["<= 30 Days"] = 0;

                                    if (this.APJ3PLReportDataTableList[num].Rows[index]["30-60"].ToString().Trim().Length != 0)
                                        dataRow[">30 and <=60 Days"] = Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["30-60"]);
                                    else
                                        dataRow[">30 and <=60 Days"] = 0;

                                    if (this.APJ3PLReportDataTableList[num].Rows[index]["60-90"].ToString().Trim().Length != 0)
                                        dataRow[">60 and <=90 Days"] = Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["60-90"]);
                                    else
                                        dataRow[">60 and <=90 Days"] = 0;

                                    if (this.APJ3PLReportDataTableList[num].Rows[index]["90-120"].ToString().Trim().Length != 0)
                                        dataRow[">90 and <=120 Days"] = Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["90-120"]);
                                    else
                                        dataRow[">90 and <=120 Days"] = 0;

                                    if (this.APJ3PLReportDataTableList[num].Rows[index]["120-150"].ToString().Trim().Length != 0)
                                        dataRow[">120 and <=150 Days"] = Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["120-150"]);
                                    else
                                        dataRow[">120 and <=150 Days"] = 0;

                                    if (this.APJ3PLReportDataTableList[num].Rows[index]["150-180"].ToString().Trim().Length != 0)
                                        dataRow[">150 and <=180 Days"] = Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["150-180"]);
                                    else
                                        dataRow[">150 and <=180 Days"] = 0;

                                    if (this.APJ3PLReportDataTableList[num].Rows[index]["180-210"].ToString().Trim().Length != 0)
                                        dataRow[">180 and <=210 Days"] = Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["180-210"]);
                                    else
                                        dataRow[">180 and <=210 Days"] = 0;

                                    if (this.APJ3PLReportDataTableList[num].Rows[index]["210-240"].ToString().Trim().Length != 0)
                                        dataRow[">210 Days"] =
                                            Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["210-240"])
                                            + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["240-270"])
                                            + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["270-300"])
                                            + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["300-330"])
                                            + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["330-360"])
                                            + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["Over 360"]);
                                    else
                                        dataRow[">210 Days"] = 0;

                                    dataRow["Qty"] =
                                        /*Convert.ToInt32(dataRow["<= 15 Days"])
                                        + Convert.ToInt32(dataRow[">15 and <=30 Days"])
                                        + Convert.ToInt32(dataRow[">30 and <=45 Days"])
                                        + Convert.ToInt32(dataRow[">45 and <=60 Days"])*/

                                        Convert.ToInt32(dataRow["<= 30 Days"])
                                        + Convert.ToInt32(dataRow[">30 and <=60 Days"])
                                        + Convert.ToInt32(dataRow[">60 and <=90 Days"])
                                        + Convert.ToInt32(dataRow[">90 and <=120 Days"])
                                        + Convert.ToInt32(dataRow[">120 and <=150 Days"])
                                        + Convert.ToInt32(dataRow[">150 and <=180 Days"])
                                        + Convert.ToInt32(dataRow[">180 and <=210 Days"])
                                        + Convert.ToInt32(dataRow[">210 Days"]);

                                    this.ConsolidationAgingDataTable.Rows.Add(dataRow);
                                    break;
                                }
                            }
                        }
                    }

                    MiscUtility.LogHistory("Done!");
                    MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
                }
                #endregion

                #region ----- Fill CCC Dragon Report raw data -----
                if (APJ3PLReportDataTableList[num].TableName.Equals("CCC"))
                {
                    MiscUtility.LogHistory(string.Format("Starting to fill 3PL CCC raw data..."));
                    MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Starting to fill 3PL CCC raw data...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));

                    for (int index = 0; index < this.APJ3PLReportDataTableList[num].Rows.Count; index++)
                    {
                        flag = false;
                        masLoc = this.APJ3PLReportDataTableList[num].Rows[index]["MAS_LOC"].ToString().Trim();

                        if (masLoc.Equals("CXM") || masLoc.Equals("EBJ") || masLoc.Equals("ESH") || masLoc.Equals("XW1"))
                        {
                            partNumber = this.APJ3PLReportDataTableList[num].Rows[index]["PART_NUM"].ToString().Trim();
                            vendorCode = this.APJ3PLReportDataTableList[num].Rows[index]["VENDOR_CODE"].ToString().Trim();
                            quantity = Convert.ToInt16(this.APJ3PLReportDataTableList[num].Rows[index]["DMI_GOODPART"].ToString().Trim());

                            for (int indey = 0; indey < this.ConsolidationAgingDataTable.Rows.Count; indey++)
                            {
                                if (masLoc.Equals(this.ConsolidationAgingDataTable.Rows[indey]["Mas Loc"])
                                        && partNumber.Equals(this.ConsolidationAgingDataTable.Rows[indey]["Part Number"]))
                                {
                                    flag = true;
                                    break;
                                }
                            }

                            if (!flag)
                            {
                                dataRow = this.ConsolidationAgingDataTable.NewRow();

                                dataRow["Date"] = currentDate;
                                dataRow["Mas Loc"] = masLoc;
                                dataRow["Vendor Code"] = vendorCode;
                                dataRow["Part Number"] = partNumber;
                                dataRow["Category"] = "";
                                /*dataRow["<= 15 Days"] = 0;
                                dataRow[">15 and <=30 Days"] = 0;
                                dataRow[">30 and <=45 Days"] = 0;
                                dataRow[">45 and <=60 Days"] = 0;*/

                                dataRow["<= 30 Days"] = 0;
                                dataRow[">30 and <=60 Days"] = 0;
                                dataRow[">60 and <=90 Days"] = 0;
                                dataRow[">90 and <=120 Days"] = 0;
                                dataRow[">120 and <=150 Days"] = 0;
                                dataRow[">150 and <=180 Days"] = 0;
                                dataRow[">180 and <=210 Days"] = 0;
                                dataRow[">210 Days"] = 0;

                                /*
                                if (ageDays <= 15)
                                    dataRow["<= 15 Days"] = quantity;

                                else if (ageDays > 15 && ageDays <= 30)
                                    dataRow[">15 and <=30 Days"] = quantity;

                                else if (ageDays > 30 && ageDays <= 45)
                                    dataRow[">30 and <=45 Days"] = quantity;

                                else if (ageDays > 45 && ageDays <= 60)
                                    dataRow[">45 and <=60 Days"] = quantity;
                                */

                                /*
                                dataRow["<= 30 Days"] =
                                    Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["<30 days"]);
                                dataRow[">30 and <=60 Days"] =
                                    Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["30-60"]);
                                dataRow[">60 and <=90 Days"] =
                                    Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["60-90"]);
                                dataRow[">90 and <=120 Days"] =
                                    Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["90-120"]);
                                dataRow[">120 and <=150 Days"] =
                                    Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["120-150"]);
                                dataRow[">150 and <=180 Days"] =
                                    Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["150-180"]);
                                dataRow[">180 and <=210 Days"] =
                                    Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["180-210"]);
                                dataRow[">210 Days"] =
                                    Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["210-240"])
                                    + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["240-270"])
                                    + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["270-300"])
                                    + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["300-330"])
                                    + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["330-360"])
                                    + Convert.ToInt32(this.APJ3PLReportDataTableList[num].Rows[index]["Over 360"]);
                                */

                                dataRow["Qty"] = quantity;
                                    /*Convert.ToInt32(dataRow["<= 15 Days"])
                                    + Convert.ToInt32(dataRow[">15 and <=30 Days"])
                                    + Convert.ToInt32(dataRow[">30 and <=45 Days"])
                                    + Convert.ToInt32(dataRow[">45 and <=60 Days"])*/

                                    /*
                                    Convert.ToInt32(dataRow["<= 30 Days"])
                                    + Convert.ToInt32(dataRow[">30 and <=60 Days"])
                                    + Convert.ToInt32(dataRow[">60 and <=90 Days"])
                                    + Convert.ToInt32(dataRow[">90 and <=120 Days"])
                                    + Convert.ToInt32(dataRow[">120 and <=150 Days"])
                                    + Convert.ToInt32(dataRow[">150 and <=180 Days"])
                                    + Convert.ToInt32(dataRow[">180 and <=210 Days"])
                                    + Convert.ToInt32(dataRow[">210 Days"]);
                                    */

                                this.ConsolidationAgingDataTable.Rows.Add(dataRow);
                            }
                        }
                    }

                    MiscUtility.LogHistory("Done!");
                    MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
                }
                #endregion
            }
        }
        #endregion

        #region ----- Fill Category for eaco part number into ConsolidationAgingDataTable -----
        private void FillCategoryAndSiteNameData()
        {
            string partNumber = string.Empty;
            string masLoc = string.Empty;

            for (int index = 0; index < this.ConsolidationAgingDataTable.Rows.Count; index++)
            {
                partNumber = this.ConsolidationAgingDataTable.Rows[index]["Part Number"].ToString();
                masLoc = this.ConsolidationAgingDataTable.Rows[index]["Mas Loc"].ToString();

                for (int indey = 0; indey < this.APJSnPPartListDataTable.Rows.Count; indey++)
                {
                    if (partNumber.Equals(this.APJSnPPartListDataTable.Rows[indey]["PartNumber"]))
                    {
                        this.ConsolidationAgingDataTable.Rows[index]["Category"] = this.APJSnPPartListDataTable.Rows[indey]["Category"];
                        this.ConsolidationAgingDataTable.Rows[index]["Description"] = this.APJSnPPartListDataTable.Rows[indey]["Description"];

                        if (this.APJSnPPartListDataTable.Rows[indey]["UnitCost"].ToString().Length != 0)
                            this.ConsolidationAgingDataTable.Rows[index]["Unit Cost"] = this.APJSnPPartListDataTable.Rows[indey]["UnitCost"];
                        else
                            this.ConsolidationAgingDataTable.Rows[index]["Unit Cost"] = 0;

                        this.ConsolidationAgingDataTable.Rows[index]["Total Cost"] =
                            Convert.ToInt32(this.ConsolidationAgingDataTable.Rows[index]["Qty"]) *
                            Convert.ToDouble(this.ConsolidationAgingDataTable.Rows[index]["Unit Cost"]);

                        break;
                    }
                }

                for (int indez = 0; indez < this.APJSnPSiteNameListDataTable.Rows.Count; indez++)
                {
                    if (masLoc.Equals(this.APJSnPSiteNameListDataTable.Rows[indez]["MasLoc"]))
                    {
                        this.ConsolidationAgingDataTable.Rows[index]["Site Name"] = this.APJSnPSiteNameListDataTable.Rows[indez]["SiteName"];
                    }
                }
            }
        }
        #endregion

        #region ----- Export consolidated aging data into excel file -----
        private void ExportConsolidationAgingReport()
        {
            string fullFileName = Path.Combine(UserSelectedValue.APJOutputFolder,
                string.Format("{0}_{1}.csv", "APJ Aging Report_", DateTime.Now.ToString("yyyy-MM-dd_HHmm")));

            MiscUtility.LogHistory(string.Format("Starting to export data into Excel file - {0}...", fullFileName));
            MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Starting to export data into Excel file - {1}...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), fullFileName));

            try
            {
                ExcelFileUtility.ExportDataIntoCSVFile(fullFileName, this.ConsolidationAgingDataTable);
                //MiscUtility.SaveFile(thirdFileName, filename);

                MiscUtility.LogHistory("Done!");
                MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Function name: <ExportConsolidationAgingReport>, Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }
        #endregion
    }
}
