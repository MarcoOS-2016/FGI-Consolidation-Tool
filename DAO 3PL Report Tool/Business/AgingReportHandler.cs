using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace DAO_3PL_Report_Tool
{
    public class AgingReportHandler
    {
        private readonly UserSelectedValue userselectedvalue = null;
        private FGIOriginalReport FGIOriginalReport = null;
        private DataTable FGIDetailsReportDatatable = new DataTable();
        private List<DataTable> AgingReportDatatableList = new List<DataTable>();
        private List<DataTable> OnhandReportDatatableList = new List<DataTable>();
        private List<DataTable> InTransitReportDataTableList = new List<DataTable>();
        
        public AgingReportHandler(UserSelectedValue userSelectValue)
        {
            this.userselectedvalue = userSelectValue;
        }

        public void Process()
        {
            SearchOriginalReports();
            LoadOriginalReports();
            BuildAgingDataTable();
            ExportReport();
        }

        private void LoadOriginalReports()
        {
            LoadAgingReport("APJ", FGIOriginalReport.APJAgingReportFile);
            LoadAgingReport("DAO", FGIOriginalReport.DAOAgingReportFile);
            LoadAgingReport("EMEA", FGIOriginalReport.EMEAAgingReportFile);

            LoadOnhandReport("APJ", FGIOriginalReport.APJOnhandReportFile);
            LoadOnhandReport("DAO", FGIOriginalReport.DAOOnhandReportFile);
            LoadOnhandReport("EMEA", FGIOriginalReport.EMEAOnhandReportFile);

            LoadInTransitReport("APJ", FGIOriginalReport.APJInTransitReportFile);
            LoadInTransitReport("DAO", FGIOriginalReport.DAOInTransitReportFile);
            LoadInTransitReport("EMEA", FGIOriginalReport.EMEAInTransitReportFile);

            LoadFGIDetailsReport(FGIOriginalReport.FGIDetailsReportFile);
        }

        #region ----- Search raw data report from specific folder
        private void SearchOriginalReports()
        {
            FGIOriginalReport = new FGIOriginalReport();

            string filePrefix = string.Empty;
            string fileName = string.Empty;            
            
            MiscUtility.LogHistory(string.Format("Starting to search report files..."));
            MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Starting to search report files...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
            
            // Search Aging report files from specified folder
            string agingReportSourceFolder = ConfigFileUtility.GetValue("AgingReportSourceFolder");
            string agingFilePrefixs = ConfigFileUtility.GetValue("AgingReportFilePrefixs");
            foreach (string agingFilePrefix in agingFilePrefixs.Split(','))
            {
                fileName = string.Format("{0}_{1}.xlsx", agingFilePrefix, userselectedvalue.WeekNumber).ToUpper();

                //if (IsExistFile(fileName, agingReportSourceFolder))
                if (File.Exists(Path.Combine(agingReportSourceFolder, fileName)))
                {
                    if (fileName.Contains("APJ"))
                        FGIOriginalReport.APJAgingReportFile = Path.Combine(agingReportSourceFolder, fileName);
                    else if (fileName.Contains("DAO"))
                        FGIOriginalReport.DAOAgingReportFile = Path.Combine(agingReportSourceFolder, fileName);
                    else if (fileName.Contains("EMEA"))
                        FGIOriginalReport.EMEAAgingReportFile = Path.Combine(agingReportSourceFolder, fileName);
                }
                else
                {
                    MiscUtility.LogHistory(string.Format("Cannot find {0} report file - {1}!", agingFilePrefix, fileName));
                    throw new Exception(string.Format("Cannot find {0} report file - {1}!", agingFilePrefix, fileName));
                }
            }

            // Search On-hand report files from specified folder
            string onhandReportSourceFolder = ConfigFileUtility.GetValue("OnhandReportSourceFolder");
            string onhandFilePrefixs = ConfigFileUtility.GetValue("OnhandReportFilePrefixs");
            foreach (string onhandFilePrefix in onhandFilePrefixs.Split(','))
            {
                fileName = string.Format("{0}_{1}.xlsx", onhandFilePrefix, userselectedvalue.WeekNumber).ToUpper();

                if (IsExistFile(fileName, onhandReportSourceFolder))
                {   
                    if (fileName.Contains("APJ"))
                        FGIOriginalReport.APJOnhandReportFile = Path.Combine(onhandReportSourceFolder, fileName);
                    else if (fileName.Contains("DAO"))
                        FGIOriginalReport.DAOOnhandReportFile = Path.Combine(onhandReportSourceFolder, fileName);
                    else if (fileName.Contains("EMEA"))
                        FGIOriginalReport.EMEAOnhandReportFile = Path.Combine(onhandReportSourceFolder, fileName);
                }
                else
                {
                    MiscUtility.LogHistory(string.Format("Cannot find {0} report file - {1}!", onhandFilePrefix, fileName));
                    throw new Exception(string.Format("Cannot find {0} report file - {1}!", onhandFilePrefix, fileName));
                }
            }

            // Search In-Transit report files from specified folder
            string intransitReportSourceFolder = ConfigFileUtility.GetValue("InTransitReportSourceFolder");
            string intransitFilePrefixs = ConfigFileUtility.GetValue("InTransitReportFilePrefixs");
            foreach (string intransitFilePrefix in intransitFilePrefixs.Split(','))
            {
                fileName = string.Format("{0}_{1}.xlsx", intransitFilePrefix, userselectedvalue.WeekNumber).ToUpper();

                if (IsExistFile(fileName, intransitReportSourceFolder))
                {
                    if (fileName.Contains("APJ"))
                        FGIOriginalReport.APJInTransitReportFile = Path.Combine(intransitReportSourceFolder, fileName);
                    else if (fileName.Contains("DAO"))
                        FGIOriginalReport.DAOInTransitReportFile = Path.Combine(intransitReportSourceFolder, fileName);
                    else if (fileName.Contains("EMEA"))
                        FGIOriginalReport.EMEAInTransitReportFile = Path.Combine(intransitReportSourceFolder, fileName);
                }
                else
                {
                    MiscUtility.LogHistory(string.Format("Cannot find {0} report file - {1}!", intransitFilePrefix, fileName));
                    throw new Exception(string.Format("Cannot find {0} report file - {1}!", intransitFilePrefix, fileName));
                }
            }

            // Search FGI Details report file from specified folder
            string detailsReportSourceFolder = ConfigFileUtility.GetValue("DetailsReportSourceFolder");
            string fgiDetailsFilePrefix = ConfigFileUtility.GetValue("FGIDetailsReportFilePrefix");
            fileName = string.Format("{0}_{1}.xlsx", fgiDetailsFilePrefix, userselectedvalue.WeekNumber).ToUpper();

            if (IsExistFile(fileName, detailsReportSourceFolder))
            {
                FGIOriginalReport.FGIDetailsReportFile = Path.Combine(detailsReportSourceFolder, fileName);
            }
            else
            {
                MiscUtility.LogHistory(string.Format("Cannot find Details report file - {0}!", fileName));
                throw new Exception(string.Format("Cannot find Details report file - {0}!", fileName));
            }

            MiscUtility.LogHistory("Done!");
            MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
        }
        #endregion

        #region ----- Load Aging reports by region -----
        private void LoadAgingReport(string regionname, string filename)
        {
            DataTable datatable = new DataTable();
           
            string sheetName = ConfigFileUtility.GetValue("AgingReportSheetName");

            MiscUtility.LogHistory(string.Format("Starting to load {0} Aging report - {1}...", regionname, filename));
            MainForm.backgroundworker.ReportProgress(0, 
                string.Format("[{0}] - Starting to load {1} Aging report - {2}...", 
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), regionname, filename));

            try
            {
                using (ExcelAccessDAO dao = new ExcelAccessDAO(filename))
                {
                    string fullSheetName = dao.IsContainSheetName(sheetName);
                    if (fullSheetName.Contains(sheetName))
                    {
                        datatable = dao.ReadAgingReportFile(sheetName);
                        datatable.TableName = regionname;
                        AgingReportDatatableList.Add(datatable);
                    }
                    else
                    {
                        MiscUtility.LogHistory(string.Format("No found the sheet name - {0} from Aging report {1}!", sheetName, filename));
                        throw new Exception(string.Format("No found the sheet name - {0} from Aging report {1}!", sheetName, filename));
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

        #region ----- Load On-hand reports by region -----
        private void LoadOnhandReport(string regionname, string filename)
        {
            DataTable datatable = new DataTable();
            
            string sheetName = ConfigFileUtility.GetValue("OnhandReportSheetName");

            MiscUtility.LogHistory(string.Format("Starting to load {0} On-hand report - {1}...", regionname, filename));
            MainForm.backgroundworker.ReportProgress(0, 
                string.Format("[{0}] - Starting to load {1} On-hand report - {2}...", 
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), regionname, filename));

            try
            {
                using (ExcelAccessDAO dao = new ExcelAccessDAO(filename))
                {
                    string fullSheetName = dao.IsContainSheetName(sheetName);

                    if (fullSheetName.Contains(sheetName))
                    {
                        datatable = dao.ReadOnhandReportFile(fullSheetName);
                        datatable.TableName = regionname;
                        OnhandReportDatatableList.Add(datatable);
                    }
                    else
                    {
                        MiscUtility.LogHistory(string.Format("No found the sheet name with the prefix - {0} from On-hand report!", sheetName));
                        throw new Exception(string.Format("No found the sheet name with the prefix - {0} from On-hand report!", sheetName));
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

        #region ----- Load In-Transit reports by region -----
        private void LoadInTransitReport(string regionname, string filename)
        {
            DataTable datatable = new DataTable();

            string sheetName = ConfigFileUtility.GetValue("InTransitReportSheetName");

            MiscUtility.LogHistory(string.Format("Starting to load {0} In-Transit report - {1}...", regionname, filename));
            MainForm.backgroundworker.ReportProgress(0,
                string.Format("[{0}] - Starting to load {1} In-Transit report - {2}...",
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), regionname, filename));

            try
            {
                using (ExcelAccessDAO dao = new ExcelAccessDAO(filename))
                {
                    string fullSheetName = dao.IsContainSheetName(sheetName);

                    if (fullSheetName.Contains(sheetName))
                    {
                        datatable = dao.ReadInTransitReportFile(fullSheetName);
                        datatable.TableName = regionname;
                        InTransitReportDataTableList.Add(datatable);
                    }
                    else
                    {
                        MiscUtility.LogHistory(string.Format("No found the sheet name with the prefix - {0} from In-Transit report!", sheetName));
                        throw new Exception(string.Format("No found the sheet name with the prefix - {0} from In-Transit report!", sheetName));
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

        #region ----- Load FGI Details report -----
        private void LoadFGIDetailsReport(string filename)
        {
            string sheetName = ConfigFileUtility.GetValue("FGIDetailsReportSheetName");

            MiscUtility.LogHistory(string.Format("Starting to load FGI Details report - {0}...", filename));
            MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Starting to load FGI Details report - {1}...",
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), filename));

            try
            {
                using (ExcelAccessDAO dao = new ExcelAccessDAO(filename))
                {
                    string fullSheetName = dao.IsContainSheetName(sheetName);
                    if (fullSheetName.Contains(sheetName))
                        FGIDetailsReportDatatable = dao.ReadProfileReportFile(sheetName);
                    else
                    {
                        MiscUtility.LogHistory(string.Format("No found the sheet name - {0} from Profile report!", sheetName));
                        throw new Exception(string.Format("No found the sheet name - {0} from Profile report!", sheetName));
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

        private void BuildAgingDataTable()
        {
            MiscUtility.LogHistory("Starting to build Consolidation Aging datatable ...");
            MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Starting to build Consolidation Aging datatable ...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));

            AddColumnsIntoDataTable();
            FillDataIntoDataTable();

            MiscUtility.LogHistory("Done!");
            MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
        }

        private void AddColumnsIntoDataTable()
        {
            try
            {
                MiscUtility.InsertNewColumn(ref FGIDetailsReportDatatable, "On_hand", "System.Int32", 18);
                MiscUtility.InsertNewColumn(ref FGIDetailsReportDatatable, "L5", "System.Int32", 19);
                MiscUtility.InsertNewColumn(ref FGIDetailsReportDatatable, "G5", "System.Int32", 20);
                MiscUtility.InsertNewColumn(ref FGIDetailsReportDatatable, "G10", "System.Int32", 21);
                MiscUtility.InsertNewColumn(ref FGIDetailsReportDatatable, "G15", "System.Int32", 22);
                MiscUtility.InsertNewColumn(ref FGIDetailsReportDatatable, "G30", "System.Int32", 23);
                MiscUtility.InsertNewColumn(ref FGIDetailsReportDatatable, "G45", "System.Int32", 24);
                MiscUtility.InsertNewColumn(ref FGIDetailsReportDatatable, "G60", "System.Int32", 25);
                MiscUtility.InsertNewColumn(ref FGIDetailsReportDatatable, "G90", "System.Int32", 26);
                MiscUtility.InsertNewColumn(ref FGIDetailsReportDatatable, "G120", "System.Int32", 27);
                MiscUtility.InsertNewColumn(ref FGIDetailsReportDatatable, "G150", "System.Int32", 28);
                MiscUtility.InsertNewColumn(ref FGIDetailsReportDatatable, "G200", "System.Int32", 29);
                MiscUtility.InsertNewColumn(ref FGIDetailsReportDatatable, "G250", "System.Int32", 30);
                MiscUtility.InsertNewColumn(ref FGIDetailsReportDatatable, "G300", "System.Int32", 31);
                MiscUtility.InsertNewColumn(ref FGIDetailsReportDatatable, "InTransit_Air", "System.Int32", 32);
                MiscUtility.InsertNewColumn(ref FGIDetailsReportDatatable, "InTransit_Ocean", "System.Int32", 33);
                MiscUtility.InsertNewColumn(ref FGIDetailsReportDatatable, "InTransit_Ground", "System.Int32", 34);
                MiscUtility.InsertNewColumn(ref FGIDetailsReportDatatable, "InTransit_Rail", "System.Int32", 35);
                MiscUtility.InsertNewColumn(ref FGIDetailsReportDatatable, "InTransit_Total", "System.Int32", 36);
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }

        private void FillDataIntoDataTable()
        {
            string fgaId = string.Empty;
            string regionName = string.Empty;

            for (int index = 0; index < FGIDetailsReportDatatable.Rows.Count; index++)
            {
                // Assigne default value to each cell
                FGIDetailsReportDatatable.Rows[index]["On_hand"] = 0;
                FGIDetailsReportDatatable.Rows[index]["L5"] = 0;
                FGIDetailsReportDatatable.Rows[index]["G5"] = 0;
                FGIDetailsReportDatatable.Rows[index]["G10"] = 0;
                FGIDetailsReportDatatable.Rows[index]["G15"] = 0;
                FGIDetailsReportDatatable.Rows[index]["G30"] = 0;
                FGIDetailsReportDatatable.Rows[index]["G45"] = 0;
                FGIDetailsReportDatatable.Rows[index]["G60"] = 0;
                FGIDetailsReportDatatable.Rows[index]["G90"] = 0;
                FGIDetailsReportDatatable.Rows[index]["G120"] = 0;
                FGIDetailsReportDatatable.Rows[index]["G150"] = 0;
                FGIDetailsReportDatatable.Rows[index]["G200"] = 0;
                FGIDetailsReportDatatable.Rows[index]["G250"] = 0;
                FGIDetailsReportDatatable.Rows[index]["G300"] = 0;
                FGIDetailsReportDatatable.Rows[index]["InTransit_Air"] = 0;
                FGIDetailsReportDatatable.Rows[index]["InTransit_Ocean"] = 0;
                FGIDetailsReportDatatable.Rows[index]["InTransit_Ground"] = 0;
                FGIDetailsReportDatatable.Rows[index]["InTransit_Rail"] = 0;
                FGIDetailsReportDatatable.Rows[index]["InTransit_Total"] = 0;

                fgaId = FGIDetailsReportDatatable.Rows[index]["fgaid"].ToString().ToUpper();

                if (FGIDetailsReportDatatable.Rows[index]["region"].ToString().ToUpper().Equals("AMERICAS")
                    || FGIDetailsReportDatatable.Rows[index]["region"].ToString().ToUpper().Equals("LATAM"))
                {
                    regionName = "DAO";
                }
                else
                {
                    regionName = FGIDetailsReportDatatable.Rows[index]["region"].ToString().ToUpper();
                }

                #region ----- Fill on-hand data into FGI Details report datatable -----
                for (int number = 0; number < OnhandReportDatatableList.Count; number++)
                {
                    if (regionName.Equals(OnhandReportDatatableList[number].TableName))
                    {
                        for (int indey = 0; indey < OnhandReportDatatableList[number].Rows.Count; indey++)
                        {
                            if (fgaId.Equals(OnhandReportDatatableList[number].Rows[indey]["FGAID"]))
                            {
                                FGIDetailsReportDatatable.Rows[index]["On_hand"] = OnhandReportDatatableList[number].Rows[indey]["On_Hand"];
                                break;
                            }
                        }
                    }
                }
                #endregion

                #region ----- Fill aging data into FGI Details report datatable -----
                string partNum = string.Empty;
                string region = string.Empty;
                int quantity = 0;
                int ageDays = 0;
                int onhandQty = 0;
                               
                for (int number = 0; number < AgingReportDatatableList.Count; number++)
                {
                    if (regionName.Equals(AgingReportDatatableList[number].TableName))
                    {
                        for (int indez = 0; indez < AgingReportDatatableList[number].Rows.Count; indez++)
                        {
                            region = AgingReportDatatableList[number].TableName;
                            partNum = AgingReportDatatableList[number].Rows[indez]["FGA"].ToString().Trim();
                            quantity = Convert.ToInt32(AgingReportDatatableList[number].Rows[indez]["Pcs"]);
                            ageDays = Convert.ToInt32(AgingReportDatatableList[number].Rows[indez]["Age"]);
                            onhandQty = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["On_hand"]);

                            if (region.Equals("EMEA") && onhandQty == 0)
                                continue;

                            if (fgaId.Equals(partNum))
                            {
                                if (ageDays > 300)
                                {
                                    FGIDetailsReportDatatable.Rows[index]["L5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["L5"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G5"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G10"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G10"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G15"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G15"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G30"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G30"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G45"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G45"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G60"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G60"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G90"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G90"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G120"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G120"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G150"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G150"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G200"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G200"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G250"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G250"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G300"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G300"]) + quantity;
                                }
                                else if (ageDays >= 251 && ageDays <= 300)
                                {
                                    FGIDetailsReportDatatable.Rows[index]["L5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["L5"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G5"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G10"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G10"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G15"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G15"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G30"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G30"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G45"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G45"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G60"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G60"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G90"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G90"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G120"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G120"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G150"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G150"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G200"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G200"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G250"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G250"]) + quantity;
                                }
                                else if (ageDays >= 201 && ageDays <= 250)
                                {
                                    FGIDetailsReportDatatable.Rows[index]["L5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["L5"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G5"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G10"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G10"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G15"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G15"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G30"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G30"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G45"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G45"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G60"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G60"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G90"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G90"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G120"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G120"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G150"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G150"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G200"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G200"]) + quantity;
                                }
                                else if (ageDays >= 151 && ageDays <= 200)
                                {
                                    FGIDetailsReportDatatable.Rows[index]["L5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["L5"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G5"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G10"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G10"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G15"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G15"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G30"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G30"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G45"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G45"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G60"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G60"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G90"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G90"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G120"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G120"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G150"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G150"]) + quantity;
                                }
                                else if (ageDays >= 121 && ageDays <= 150)
                                {
                                    FGIDetailsReportDatatable.Rows[index]["L5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["L5"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G5"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G10"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G10"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G15"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G15"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G30"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G30"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G45"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G45"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G60"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G60"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G90"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G90"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G120"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G120"]) + quantity;
                                }
                                else if (ageDays >= 91 && ageDays <= 120)
                                {
                                    FGIDetailsReportDatatable.Rows[index]["L5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["L5"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G5"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G10"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G10"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G15"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G15"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G30"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G30"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G45"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G45"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G60"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G60"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G90"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G90"]) + quantity;
                                }
                                else if (ageDays >= 61 && ageDays <= 90)
                                {
                                    FGIDetailsReportDatatable.Rows[index]["L5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["L5"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G5"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G10"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G10"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G15"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G15"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G30"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G30"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G45"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G45"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G60"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G60"]) + quantity;
                                }
                                else if (ageDays >= 46 && ageDays <= 60)
                                {
                                    FGIDetailsReportDatatable.Rows[index]["L5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["L5"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G5"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G10"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G10"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G15"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G15"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G30"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G30"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G45"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G45"]) + quantity;
                                }
                                else if (ageDays >= 31 && ageDays <= 45)
                                {
                                    FGIDetailsReportDatatable.Rows[index]["L5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["L5"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G5"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G10"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G10"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G15"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G15"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G30"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G30"]) + quantity;
                                }
                                else if (ageDays >= 16 && ageDays <= 30)
                                {
                                    FGIDetailsReportDatatable.Rows[index]["L5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["L5"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G5"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G10"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G10"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G15"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G15"]) + quantity;
                                }
                                else if (ageDays >= 11 && ageDays <= 15)
                                {
                                    FGIDetailsReportDatatable.Rows[index]["L5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["L5"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G5"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G10"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G10"]) + quantity;
                                }
                                else if (ageDays >= 6 && ageDays <= 10)
                                {
                                    FGIDetailsReportDatatable.Rows[index]["L5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["L5"]) + quantity;
                                    FGIDetailsReportDatatable.Rows[index]["G5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["G5"]) + quantity;
                                }
                                else if (ageDays <= 5)
                                {
                                    FGIDetailsReportDatatable.Rows[index]["L5"] = Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["L5"]) + quantity;
                                }
                            }
                        }
                    }
                }
                #endregion

                #region ----- Fill In-Transit data into FGI Details report datatable -----                
                string shipMode = string.Empty;
                int qty = 0;

                for (int number = 0; number < InTransitReportDataTableList.Count; number++)
                {
                    if (regionName.Equals(AgingReportDatatableList[number].TableName))
                    {
                        for (int indez = 0; indez < InTransitReportDataTableList[number].Rows.Count; indez++)
                        {
                            if (InTransitReportDataTableList[number].Rows[indez]["FGAID"].ToString().Trim().Length == 0)
                                continue;

                            partNum = InTransitReportDataTableList[number].Rows[indez]["FGAID"].ToString().Trim();
                            qty = Convert.ToInt32(InTransitReportDataTableList[number].Rows[indez]["Quantity"]);

                            if (fgaId.Equals(partNum))
                            {
                                if (InTransitReportDataTableList[number].Rows[indez]["Shipmode"].ToString().ToUpper().Equals("OCEAN SHUTTLE TRUCK")
                                || InTransitReportDataTableList[number].Rows[indez]["Shipmode"].ToString().ToUpper().Equals("OCEAN CONSOLIDATED"))
                                    shipMode = "OCEAN";
                                else if (InTransitReportDataTableList[number].Rows[indez]["Shipmode"].ToString().ToUpper().Equals("OCEAN SHUTTLE RAIL"))
                                    shipMode = "RAIL";
                                else
                                    shipMode = InTransitReportDataTableList[number].Rows[indez]["Shipmode"].ToString().ToUpper();

                                if (shipMode.Equals("AIR"))
                                {
                                    FGIDetailsReportDatatable.Rows[index]["InTransit_Air"] =
                                        Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["InTransit_Air"]) + qty;
                                    FGIDetailsReportDatatable.Rows[index]["InTransit_Total"] =
                                        Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["InTransit_Total"]) + qty;
                                }
                                else if (shipMode.Equals("OCEAN"))
                                {
                                    FGIDetailsReportDatatable.Rows[index]["InTransit_Ocean"] =
                                        Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["InTransit_Ocean"]) + qty;
                                    FGIDetailsReportDatatable.Rows[index]["InTransit_Total"] =
                                        Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["InTransit_Total"]) + qty;
                                }
                                else if (shipMode.Equals("GROUND"))
                                {
                                    FGIDetailsReportDatatable.Rows[index]["InTransit_Ground"] =
                                        Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["InTransit_Ground"]) + qty;
                                    FGIDetailsReportDatatable.Rows[index]["InTransit_Total"] =
                                        Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["InTransit_Total"]) + qty;
                                }
                                else if (shipMode.Equals("RAIL"))
                                {
                                    FGIDetailsReportDatatable.Rows[index]["InTransit_Rail"] =
                                        Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["InTransit_Rail"]) + qty;
                                    FGIDetailsReportDatatable.Rows[index]["InTransit_Total"] =
                                        Convert.ToInt32(FGIDetailsReportDatatable.Rows[index]["InTransit_Total"]) + qty;
                                }
                            }
                        }
                    }
                }
                #endregion
            }
        }

        private void ExportReport()
        {
            //string fileName = Path.Combine(userselectedvalue.AgingReportOutputFolder,
            //    string.Format("{0}_{1}.csv", "Consolidation_Aging", DateTime.Now.ToString("yyyy-MM-dd_HHmm")));

            string consolidationAgingFilePrefix = ConfigFileUtility.GetValue("ConsolidationAgingFilePrefix");
            string fileName = Path.Combine(userselectedvalue.AgingReportOutputFolder,
                string.Format("{0}_{1}.csv", consolidationAgingFilePrefix, userselectedvalue.WeekNumber));

            MiscUtility.LogHistory(string.Format("Starting to export data into Excel file - {0}...", fileName));
            MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Starting to export data into Excel file - {1}...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), fileName));

            try
            {
                ExcelFileUtility.ExportDataIntoCSVFile(fileName, FGIDetailsReportDatatable);
                //MiscUtility.SaveFile(thirdFileName, filename);

                MiscUtility.LogHistory("Done!");
                MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Function name: <ExportToExcelFile>, Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }

        private bool IsExistFile(string filename, string sourcefolder)
        {
            bool isExist = false;
            DirectoryInfo dir = null;

            try
            {
                MiscUtility.LogHistory(string.Format("Starting to search file - {0}", filename));
                MainForm.backgroundworker.ReportProgress(0,
                    string.Format("[{0}] - Starting to search file - {1}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), filename));

                dir = new DirectoryInfo(sourcefolder);
                foreach (FileInfo fi in dir.GetFiles())
                {
                    if (fi.Name.ToUpper().Equals(filename))
                    {
                        isExist = true;
                        break;
                    }
                }

                return isExist;
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Function name: <IsExistFile>, Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }
    }

}
